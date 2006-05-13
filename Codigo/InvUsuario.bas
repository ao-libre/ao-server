Attribute VB_Name = "InvUsuario"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(UserIndex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

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
Next

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UserList(UserIndex).Pos.Map = 37 Then
    
    Dim DeDonde As WorldPos
    
    Select Case UCase$(UserList(UserIndex).Hogar)
        Case "LINDOS" 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case "ULLATHORPE"
            DeDonde = Ullathorpe
        Case "BANDERBILL"
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
       
    Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
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

UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
On Error GoTo errhandler

If Cantidad > 100000 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon
        If Cantidad > 39999 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.Map
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
        
        Do While (Cantidad > 0) And (UserList(UserIndex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name, False)
            
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)

'Quita un objeto
UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(UserIndex).Invent.Object(Slot).Amount <= 0 Then
    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    UserList(UserIndex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

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

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(UserIndex, Slot)
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        If num + MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        End If
        
        Obj.Amount = num
        
        Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!" & FONTTYPE_TALK)
        End If
        If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡ATENCION!! ¡¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!!" & FONTTYPE_TALK)
        End If
        
        If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).name, False)
  Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
   Else
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    End If
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).OBJInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount + Obj.Amount
    Else
        MapData(Map, X, Y).OBJInfo = Obj
        
        If sndRoute = SendTarget.ToMap Then
            Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
        Else
            Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
        End If
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
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
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes cargar mas objetos." & FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name, False)
        End If
        
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay nada aqui." & FONTTYPE_INFO)
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
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otHerramientas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
End Select

Call SendUserStatsBox(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(UserIndex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(UserIndex) Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.ArmadaReal = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(UserIndex) Then
        FaccionPuedeUsarItem = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
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
                        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
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
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_SACARARMA)
        
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
       Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
    
    Case eOBJType.otHerramientas
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.HerramientaEqpSlot = Slot
                
       Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
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
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
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
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
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
                UserList(UserIndex).CharMimetizado.Body = Obj.Ropaje
            Else
                UserList(UserIndex).Char.Body = Obj.Ropaje
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            UserList(UserIndex).flags.Desnudo = 0
            

        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase,genero o raza no puede usar este objeto." & FONTTYPE_INFO)
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
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
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
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
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
                     Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
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
                 
                 Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             End If
         Else
             Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

'Verifica si la raza puede usar la ropa
If UserList(UserIndex).Raza = "Humano" Or _
   UserList(UserIndex).Raza = "Elfo" Or _
   UserList(UserIndex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(UserIndex).Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
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
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + Obj.MinHam
        If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then _
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).flags.Hambre = 0
        Call EnviarHambreYsed(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call SendUserStatsBox(UserIndex)
        
    Case eOBJType.otWeapon
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
        End If
        
        If ObjData(ObjIndex).proyectil = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Proyectiles)
        Else
            If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
            
            '¿El target-objeto es leña?
            If UserList(UserIndex).flags.TargetObj = Leña Then
                If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, _
                         UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                End If
            End If
        End If
    
    Case eOBJType.otPociones
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteAtacar(UserIndex, False) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Debes esperar unos momentos para tomar otra pocion!!" & FONTTYPE_INFO)
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
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
            Case 2 'Modif la fuerza
                UserList(UserIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then _
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)
                If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
            Case 6  ' Pocion Negra
                If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDie(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Sientes un gran mareo y pierdes el conocimiento." & FONTTYPE_FIGHT)
                End If
       End Select
       Call SendUserStatsBox(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call EnviarHambreYsed(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_BEBER)
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
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
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
    
        Case eOBJType.otBotellaVacia
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            Call UpdateUserInv(False, UserIndex, Slot)
    
        Case eOBJType.otBotellaLlena
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + Obj.MinSed
            If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
                UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
            UserList(UserIndex).flags.Sed = 0
            Call EnviarHambreYsed(UserIndex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
            
            
        Case eOBJType.otHerramientas
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(UserIndex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                UserList(UserIndex).Reputacion.PlebeRep = MAXREP
            
            Select Case ObjIndex
                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SFC")

            End Select
        
        Case eOBJType.otPergaminos
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Hambre = 0 And _
               UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
            End If
       
       Case eOBJType.otMinerales
           If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
           End If
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & FundirMetal)
       
       Case eOBJType.otInstrumentos
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Obj.Snd1)
       
       Case eOBJType.otBarcos
    'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(UserIndex).Stats.ELV < 25 Then
            If UCase$(UserList(UserIndex).Clase) <> "PESCADOR" And UCase$(UserList(UserIndex).Clase) <> "PIRATA" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para recorrer los mares debes ser nivel 25 o superior." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True) Or _
            LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)) And _
            UserList(UserIndex).flags.Navegando = 0) _
            Or UserList(UserIndex).flags.Navegando = 1 Then
           Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)
        End If
           
End Select

'Actualiza
'Call SendUserStatsBox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(UserIndex).Clase) Then
        If ObjData(ArmasHerrero(i)).OBJType = eOBJType.otWeapon Then
        '[DnG!]
            cad$ = cad$ & ObjData(ArmasHerrero(i)).name & " (" & ObjData(ArmasHerrero(i)).LingH & "-" & ObjData(ArmasHerrero(i)).LingP & "-" & ObjData(ArmasHerrero(i)).LingO & ")" & "," & ArmasHerrero(i) & ","
        '[/DnG!]
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(SendTarget.ToIndex, UserIndex, 0, "LAH" & cad$)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) / ModCarpinteria(UserList(UserIndex).Clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(SendTarget.ToIndex, UserIndex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).Clase) Then
        '[DnG!]
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).name & " (" & ObjData(ArmadurasHerrero(i)).LingH & "-" & ObjData(ArmadurasHerrero(i)).LingP & "-" & ObjData(ArmadurasHerrero(i)).LingO & ")" & "," & ArmadurasHerrero(i) & ","
        '[/DnG!]
    End If
Next i

Call SendData(SendTarget.ToIndex, UserIndex, 0, "LAR" & cad$)

End Sub


                   

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(UserIndex)
Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)

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
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
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

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
