Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Sub IniciarDeposito(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    'Hacemos un Update del inventario del usuario
    Call UpdateBanUserInv(True, Userindex, 0)
    Call WriteBankInit(Userindex)

    UserList(Userindex).flags.Comerciando = True

ErrHandler:

End Sub

Sub SendBanObj(Userindex As Integer, Slot As Byte, Object As UserObj)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(Userindex).BancoInvent.Object(Slot) = Object

    Call WriteChangeBankSlot(Userindex, Slot)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, _
                     ByVal Userindex As Integer, _
                     ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NullObj As UserObj

    Dim LoopC   As Byte

    With UserList(Userindex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .BancoInvent.Object(Slot).ObjIndex > 0 Then
                Call SendBanObj(Userindex, Slot, .BancoInvent.Object(Slot))
            Else
                Call SendBanObj(Userindex, Slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
                    Call SendBanObj(Userindex, LoopC, .BancoInvent.Object(LoopC))
                Else
                    Call SendBanObj(Userindex, LoopC, NullObj)

                End If

            Next LoopC

        End If

    End With

End Sub

Sub UserRetiraItem(ByVal Userindex As Integer, _
                   ByVal BankSlot As Integer, _
                   ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    Dim InvSlot As Integer

    If Cantidad < 1 Then Exit Sub
    
    Call WriteUpdateUserStats(Userindex)

    If UserList(Userindex).BancoInvent.Object(BankSlot).Amount > 0 Then
    
        If Cantidad > UserList(Userindex).BancoInvent.Object(BankSlot).Amount Then Cantidad = UserList(Userindex).BancoInvent.Object(BankSlot).Amount
            
        ObjIndex = UserList(Userindex).BancoInvent.Object(BankSlot).ObjIndex
        
        'Agregamos el obj que compro al inventario
        InvSlot = UserReciveObj(Userindex, BankSlot, Cantidad)
        
        If InvSlot > 0 Then
            If ObjData(ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(Userindex).Name & " retiro " & Cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
            End If
            
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, Userindex, InvSlot)
            'Actualizamos el banco
            Call UpdateBanUserInv(False, Userindex, BankSlot)
        End If

    End If

ErrHandler:

End Sub

Function UserReciveObj(ByVal Userindex As Integer, _
                  ByVal InvSlot As Integer, _
                  ByVal Cantidad As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Slot As Integer

    Dim obji As Integer

    With UserList(Userindex)

        If .BancoInvent.Object(InvSlot).Amount <= 0 Then Exit Function
    
        obji = .BancoInvent.Object(InvSlot).ObjIndex
    
        'Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Exit Do

            End If

        Loop
    
        'Sino se fija por un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(Userindex, "No podes tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            Loop
            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        'Mete el obj en el slot
        If .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = obji
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + Cantidad
        
            Call QuitarBancoInvItem(Userindex, InvSlot, Cantidad)

            UserReciveObj = Slot
        Else
            Call WriteConsoleMsg(Userindex, "No podes tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Function

Sub QuitarBancoInvItem(ByVal Userindex As Integer, _
                       ByVal Slot As Byte, _
                       ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ObjIndex As Integer

    With UserList(Userindex)
        ObjIndex = .BancoInvent.Object(Slot).ObjIndex

        'Quita un Obj

        .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - Cantidad
    
        If .BancoInvent.Object(Slot).Amount <= 0 Then
            .BancoInvent.NroItems = .BancoInvent.NroItems - 1
            .BancoInvent.Object(Slot).ObjIndex = 0
            .BancoInvent.Object(Slot).Amount = 0

        End If

    End With
    
End Sub

Sub UserDepositaItem(ByVal Userindex As Integer, _
                     ByVal InvSlot As Integer, _
                     ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    '06/04/2020: FrankoH298 - No podemos vender monturas en uso.
    '***************************************************

    Dim ObjIndex As Integer
    Dim BankSlot As Integer
    With UserList(Userindex)
        If .flags.Equitando = 1 Then
            If .Invent.MonturaEqpSlot = InvSlot Then
                Call WriteConsoleMsg(Userindex, "No podes depositar tu montura mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        End If
        If .Invent.Object(InvSlot).Amount > 0 And Cantidad > 0 Then
        
            If Cantidad > .Invent.Object(InvSlot).Amount Then Cantidad = .Invent.Object(InvSlot).Amount
            
            ObjIndex = .Invent.Object(InvSlot).ObjIndex
            
            'Agregamos el obj que deposita al banco
            BankSlot = UserDejaObj(Userindex, InvSlot, Cantidad)
            
            If BankSlot > 0 Then
                If ObjData(ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " deposito " & Cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
                End If
                
                'Actualizamos el inventario del usuario
                Call UpdateUserInv(False, Userindex, InvSlot)
                
                'Actualizamos el inventario del banco
                Call UpdateBanUserInv(False, Userindex, BankSlot)
            End If
    
        End If
    End With
End Sub

Function UserDejaObj(ByVal Userindex As Integer, _
                ByVal InvSlot As Integer, _
                ByVal Cantidad As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Slot As Integer

    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Function
    
    With UserList(Userindex)
        obji = .Invent.Object(InvSlot).ObjIndex
        
        'Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .BancoInvent.Object(Slot).ObjIndex = obji And .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do

            End If

        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1

            Do Until .BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteConsoleMsg(Userindex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1

        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
            If .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
                
                'Menor que MAX_INV_OBJS
                .BancoInvent.Object(Slot).ObjIndex = obji
                .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + Cantidad
                
                Call QuitarUserInvItem(Userindex, InvSlot, Cantidad)
                
                UserDejaObj = Slot
            Else
                Call WriteConsoleMsg(Userindex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Function

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(Userindex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(Userindex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(Userindex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(Userindex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(Userindex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next

End Sub
