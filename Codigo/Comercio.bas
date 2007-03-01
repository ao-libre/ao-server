Attribute VB_Name = "Comercio"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


Function UserCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal NpcIndex As Integer, ByVal Cantidad As Integer) As Boolean
On Error GoTo errorh
    Dim Descuento As String
    Dim unidad As Long, monto As Long
    Dim Slot As Integer
    Dim obji As Integer
    Dim Encontre As Boolean
    
    UserCompraObj = False
    
    
    
    If (Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(ObjIndex).amount <= 0) Then Exit Function
    
    obji = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(ObjIndex).ObjIndex
    
    
    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And _
       UserList(UserIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS
        
        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            
            If Slot > MAX_INVENTORY_SLOTS Then
                Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
    End If
    
    'desde aca para abajo se realiza la transaccion
    UserCompraObj = True
    'Mete el obj en el slot
    If UserList(UserIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
        UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + Cantidad
        
        'Le sustraemos el valor en oro del obj comprado
        Descuento = UserList(UserIndex).flags.Descuento
        If Descuento = 0 Then Descuento = 1 'evitamos dividir por 0!
        unidad = ((ObjData(Npclist(NpcIndex).Invent.Object(ObjIndex).ObjIndex).Valor) / Descuento)
        monto = unidad * Cantidad
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - monto
        
        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(UserIndex, Comerciar)
        
        If ObjData(obji).OBJType = eOBJType.otLlaves Then Call logVentaCasa(UserList(UserIndex).name & " compro " & ObjData(obji).name)
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(ObjIndex), Cantidad)
    Else
        Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
Exit Function

errorh:
Call LogError("Error en USERCOMPRAOBJ. " & Err.description)
End Function


Sub NpcCompraObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo errorh
    Dim Slot As Integer
    Dim obji As Integer
    Dim NpcIndex As Integer
    Dim infla As Long
    Dim monto As Long
          
    If Cantidad < 1 Then Exit Sub
    
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
    
    If ObjData(obji).Newbie = 1 Then
        Call WriteConsoleMsg(UserIndex, "No comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera Then
        '¿Son los items con los que comercia el npc?
        If Npclist(NpcIndex).TipoItems <> ObjData(obji).OBJType Then
            Call WriteConsoleMsg(UserIndex, "El npc no esta interesado en comprar ese objeto.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
    End If
    
    If obji = iORO Then
        Call WriteConsoleMsg(UserIndex, "El npc no esta interesado en comprar ese objeto.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    
    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until (Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji _
      And Npclist(NpcIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS)
        
        Slot = Slot + 1
        
        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop
    
    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then Exit Do
        Loop
        If Slot <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    End If
    
    If Slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji
        If Npclist(NpcIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            Npclist(NpcIndex).Invent.Object(Slot).amount = Npclist(NpcIndex).Invent.Object(Slot).amount + Cantidad
        Else
            Npclist(NpcIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
        End If
    End If
    
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
    'Le sumamos al user el valor en oro del obj vendido
    monto = ((ObjData(obji).Valor \ 3) * Cantidad)
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + monto
    If UserList(UserIndex).Stats.GLD > MAXORO Then _
        UserList(UserIndex).Stats.GLD = MAXORO
    
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(UserIndex, Comerciar)
Exit Sub

errorh:
    Call LogError("Error en NPCCOMPRAOBJ. " & Err.description)
End Sub

Sub IniciarCOmercioNPC(ByVal UserIndex As Integer)
On Error GoTo errhandler
    'Mandamos el Inventario
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    'Hacemos un Update del inventario del usuario
    Call UpdateUserInv(True, UserIndex, 0)
    'Atcualizamos el dinero
    Call SendUserStatsBox(UserIndex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)
Exit Sub

errhandler:
    Dim str As String
    str = "Error en IniciarComercioNPC. UI=" & UserIndex
    If UserIndex > 0 Then
        str = str & ".Nombre: " & UserList(UserIndex).name & " IP:" & UserList(UserIndex).ip & " comerciando con "
        If UserList(UserIndex).flags.TargetNPC > 0 Then
            str = str & Npclist(UserList(UserIndex).flags.TargetNPC).name
        Else
            str = str & "<NPCINDEX 0>"
        End If
    Else
        str = str & "<USERINDEX 0>"
    End If
End Sub

Sub NPCVentaItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)
'listindex+1, cantidad
On Error GoTo errhandler

    Dim val As Long
    Dim desc As String
    
    If Cantidad < 1 Then Exit Sub
    
    'NPC VENDE UN OBJ A UN USUARIO
    Call SendUserStatsBox(UserIndex)
    
    If i > MAX_INVENTORY_SLOTS Then
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Posible intento de romper el sistema de comercio. Usuario: " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_WARNING))
        Exit Sub
    End If
    
    If Cantidad > MAX_INVENTORY_OBJS Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
        Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
        UserList(UserIndex).flags.Ban = 1
        Call WriteErrorMsg(UserIndex, "Has sido baneado por el sistema anti cheats")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    'Calculamos el valor unitario
    desc = Descuento(UserIndex)
    If desc = 0 Then desc = 1 'evitamos dividir por 0!
    val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / desc
    
    If val = 0 Then val = 1 'Evita que un objeto valga 0
    
    If UserList(UserIndex).Stats.GLD >= (val * Cantidad) Then
        If Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).amount > 0 Then
            If Cantidad > Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(i).amount
            'Agregamos el obj que compro al inventario
            If Not UserCompraObj(UserIndex, CInt(i), UserList(UserIndex).flags.TargetNPC, Cantidad) Then
                Call WriteConsoleMsg(UserIndex, "No puedes comprar este ítem.", FontTypeNames.FONTTYPE_INFO)
            End If
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el oro
            Call SendUserStatsBox(UserIndex)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call UpdateVentanaComercio(UserIndex)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No tenes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
    End If
Exit Sub

errhandler:
    Call LogError("Error en comprar item: " & Err.description)
End Sub

Sub NPCCompraItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler
    Dim NpcIndex As Integer
    
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    'Si es una armadura faccionaria vemos que la está intentando vender al sastre
    If ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Real = 1 Then
        If Npclist(NpcIndex).name <> "SR" Then
            Call WriteConsoleMsg(UserIndex, "Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres", FontTypeNames.FONTTYPE_WARNING)
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        End If
    ElseIf ObjData(UserList(UserIndex).Invent.Object(Item).ObjIndex).Caos = 1 Then
        If Npclist(NpcIndex).name <> "SC" Then
            Call WriteConsoleMsg(UserIndex, "Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres", FontTypeNames.FONTTYPE_WARNING)
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(UserIndex)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        End If
    ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
        Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    'NPC COMPRA UN OBJ A UN USUARIO
    Call SendUserStatsBox(UserIndex)
   
    If UserList(UserIndex).Invent.Object(Item).amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).amount
        'Agregamos el obj que compro al inventario
        Call NpcCompraObj(UserIndex, CInt(Item), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        'Actualizamos el oro
        Call SendUserStatsBox(UserIndex)
        
        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
        'Actualizamos la ventana de comercio
        Call UpdateVentanaComercio(UserIndex)
    End If
Exit Sub

errhandler:
    Call LogError("Error en vender item: " & Err.description)
End Sub

Sub UpdateVentanaComercio(ByVal UserIndex As Integer)
    Call WriteTradeOK(UserIndex)
End Sub

Function Descuento(ByVal UserIndex As Integer) As Single
    'Calcula el descuento al comerciar
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
    UserList(UserIndex).flags.Descuento = Descuento
End Function

Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i As Integer
    Dim desc As String
    Dim val As Long

    desc = Descuento(UserIndex)
    If desc = 0 Then desc = 1 'evitamos dividir por 0!
    
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex > 0 Then
            Dim thisObj As Obj
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(i).ObjIndex
            thisObj.amount = Npclist(NpcIndex).Invent.Object(i).amount
            'Calculamos el porc de inflacion del npc
            val = (ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).Valor) / desc
            
            Call WriteChangeNPCInventorySlot(UserIndex, thisObj, val)
        Else
             Dim DummyObj As Obj
             Call WriteChangeNPCInventorySlot(UserIndex, DummyObj, 0)
        End If
    Next i
End Sub
