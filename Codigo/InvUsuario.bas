Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UCase$(MapInfo(UserList(UserIndex).Pos.map).Restringir) = "NEWBIE" Then
    
    Dim DeDonde As WorldPos
    
    Select Case UserList(UserIndex).Hogar
        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case eCiudad.cUllathorpe
            DeDonde = Ullathorpe
        Case eCiudad.cBanderbill
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
    
    Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If Cantidad > 50000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.map
            For j = UserList(UserIndex).Pos.X - 10 To UserList(UserIndex).Pos.X + 10
                For k = UserList(UserIndex).Pos.Y - 10 To UserList(UserIndex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(UserIndex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.amount
            Else
                MiObj.amount = Cantidad
                Cantidad = Cantidad - MiObj.amount
            End If

            MiObj.ObjIndex = iORO
            
            If EsGM(UserIndex) Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            Dim AuxPos As WorldPos
            
            If UserList(UserIndex).clase = eClass.Pirat And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .amount = .amount - Cantidad
        '¿Quedan mas?
        If .amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).amount Then num = UserList(UserIndex).Invent.Object(Slot).amount
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If num + MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        End If
        
        Obj.amount = num
        
        Call MakeObj(Obj, map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call WriteConsoleMsg(UserIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
        End If
        
        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).name)
        
        'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Obj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " tiró al piso " & Obj.amount & " " & ObjData(Obj.ObjIndex).name)
        ElseIf Obj.amount = 1000 Then 'Es mucha cantidad?
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Obj.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " tiró del piso " & Obj.amount & " " & ObjData(Obj.ObjIndex).name)
            End If
        End If
  Else
    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount - num

If MapData(map, X, Y).ObjInfo.amount <= 0 Then
    MapData(map, X, Y).ObjInfo.ObjIndex = 0
    MapData(map, X, Y).ObjInfo.amount = 0
    
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount + Obj.amount
    Else
        MapData(map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
Else
   UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.amount = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            
            'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
            ElseIf MiObj.amount = 1000 Then 'Es mucha cantidad?
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                    Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
                End If
            End If
            
        End If
        
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otAnillo
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not criminal(UserIndex) Then
        FaccionPuedeUsarItem = esArmada(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If criminal(UserIndex) Then
        FaccionPuedeUsarItem = esCaos(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
                Else
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
            Else
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otAnillo
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.body = Obj.Ropaje
            Else
                UserList(UserIndex).Char.body = Obj.Ropaje
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            UserList(UserIndex).flags.Desnudo = 0
            

        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otCASCO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.CascoEqpSlot = Slot
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(UserIndex, Slot)
                 If UserList(UserIndex).flags.Mimetizado = 1 Then
                     UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(UserIndex).Invent.Object(Slot).Equipped = 1
             UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
             UserList(UserIndex).Invent.EscudoEqpSlot = Slot
             
             If UserList(UserIndex).flags.Mimetizado = 1 Then
                 UserList(UserIndex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                 
                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             End If
         Else
             Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

'Verifica si la raza puede usar la ropa
If UserList(UserIndex).raza = eRaza.Humano Or _
   UserList(UserIndex).raza = eRaza.Elfo Or _
   UserList(UserIndex).raza = eRaza.Drow Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If

'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
If (UserList(UserIndex).raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
    CheckRazaUsaRopa = False
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'*************************************************

Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(UserIndex).Invent.Object(Slot).amount = 0 Then Exit Sub

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        If Not UserList(UserIndex).flags.ModoCombate Then
            Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + Obj.MinHam
        If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then _
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).amount
        UserList(UserIndex).Invent.Object(Slot).amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call WriteConsoleMsg(UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not UserList(UserIndex).flags.ModoCombate Then
                Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        Else
            If UserList(UserIndex).flags.TargetObj = Leña Then
                If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, _
                         UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                End If
            End If
        End If
        
        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjIndex
            Case CAÑA_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
            Case HACHA_LEÑADOR
                Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
            Case PIQUETE_MINERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteAtacar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(UserIndex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
            Case 2 'Modif la fuerza
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then _
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)
                If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Case 6  ' Pocion Negra
                If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
       End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
            Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.MaxMAN > 0 Then
            If UserList(UserIndex).flags.Hambre = 0 And _
                UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If UserList(UserIndex).flags.Muerto = 1 Then
             Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.map, PrepareMessagePlayWave(Obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.map, PrepareMessagePlayWave(Obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
       
    Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(UserIndex).Stats.ELV < 25 Then
            If UserList(UserIndex).clase <> eClass.Fisher And UserList(UserIndex).clase <> eClass.Pirat Then
                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If UserList(UserIndex).Stats.ELV < 20 Then
                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        
        If ((LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True, False)) _
                And UserList(UserIndex).flags.Navegando = 0) _
                Or UserList(UserIndex).flags.Navegando = 1 Then
            Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
End Select

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(UserIndex)

Dim Cantidad As Long
Cantidad = UserList(UserIndex).Stats.GLD - CLng(UserList(UserIndex).Stats.ELV) * 10000

If Cantidad > 0 Then _
    Call TirarOro(Cantidad, UserIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.amount = UserList(UserIndex).Invent.Object(i).amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(UserIndex).clase = eClass.Pirat And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
