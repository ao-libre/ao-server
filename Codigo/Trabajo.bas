Attribute VB_Name = "Trabajo"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Private Const GASTO_ENERGIA_TRABAJADOR    As Byte = 2

Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6

Public Sub DoPermanecerOculto(ByVal Userindex As Integer, ByVal DeltaTick As Single)

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo errHandler

    With UserList(Userindex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - DeltaTick

        If .Counters.TiempoOculto <= 0 Then
            If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub

                End If

            End If

            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(Userindex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoPermanecerOculto")

End Sub

Public Sub DoOcultarse(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la formula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo errHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(Userindex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .Clase = eClass.Bandit Then
                .Counters.TiempoOculto = Int(Suerte / 2)
            Else
                .Counters.TiempoOculto = Suerte

            End If
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(Userindex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(Userindex, "Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)

            End If
            
            Call SubirSkill(Userindex, eSkill.Ocultarse, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(Userindex, "No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If

            '[/CDT]
            
            Call SubirSkill(Userindex, eSkill.Ocultarse, False)

        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal Userindex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/07/2020: Jopi
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la ultima barca que equipo era el galeon no explotaba(Y capaz no la tenia equipada :P).
'12/01/2020: Recox - Se refactorizo un poco para reutilizar con monturas.
'09/07/2020: Jopi - Dejamos de usar los modificadores de barca por clase y refactorizo los chequeos por clase.
'***************************************************

    With UserList(Userindex)
    
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes navegar mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If EsGalera(Barco) Then
            
            If .Clase <> eClass.Assasin And _
                .Clase <> eClass.Pirat And _
                .Clase <> eClass.Bandit And _
                .Clase <> eClass.Cleric And _
                .Clase <> eClass.Thief And _
                .Clase <> eClass.Paladin Then
            
                Call WriteConsoleMsg(Userindex, "Solo los Piratas, Asesinos, Bandidos, Clerigos, Bandidos y Paladines pueden usar Galera!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
            End If
            
        End If
        
        If EsGaleon(Barco) Then
            
            If .Clase <> eClass.Thief And .Clase <> eClass.Pirat Then
            
                Call WriteConsoleMsg(Userindex, "Solo los Ladrones y Piratas pueden usar Galeon!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
            End If
            
        End If
        
        ' Acordate que el Trabajador solo necesita 60 de Navegacion para usar barca!
        Dim SkillNecesario As Byte
        SkillNecesario = IIf(.Clase = eClass.Worker, 60, Barco.MinSkill)
        
        ' Tiene el skill necesario?
        If .Stats.UserSkills(eSkill.Navegacion) < SkillNecesario Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.BarcoSlot = Slot
            
            .Char.Head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(Userindex)
                Call SetVisibleStateForUserAfterNavigateOrEquitate(Userindex)
                
            ' Esta muerto
            Else
                .Char.body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
        ' Estaba navegando
        Else
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
        
            ' No esta muerto
            If .flags.Muerto = 0 Then
                If .flags.Mimetizado = 0 Then
                    .Char.Head = .OrigChar.Head
                End If
                
                Call SetEquipmentOnCharAfterNavigateOrEquitate(Userindex)
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.invisible = 1 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, True)
                End If
                
            ' Esta muerto
            Else
                .Char.body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

            End If
            
            ' Termina de navegar
            .flags.Navegando = 0

        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

    End With
    
    Call WriteNavigateToggle(Userindex)

End Sub

Public Sub FundirMineral(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(Userindex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) / ModFundicion(.Clase) Then
                Call DoLingotes(Userindex)
            Else
                Call WriteConsoleMsg(Userindex, "No tienes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Public Sub FundirArmas(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(Userindex)

        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                If ObjData(.flags.TargetObjInvIndex).SkHerreria <= .Stats.UserSkills(eSkill.Herreria) / ModHerreriA(.Clase) Then
                    Call DoFundir(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes los conocimientos suficientes en herreria para fundir este objeto.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en FundirArmas. Error " & Err.Number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
    '***************************************************

    Dim i     As Integer

    Dim Total As Long

    For i = 1 To UserList(Userindex).CurrentInventorySlots

        If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(Userindex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Integer, _
                         ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************

    Dim i As Integer

    For i = 1 To UserList(Userindex).CurrentInventorySlots

        With UserList(Userindex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(Userindex, i)
                
                .Amount = .Amount - cant

                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, Userindex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, _
                            ByVal ItemIndex As Integer, _
                            ByVal CantidadItems As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, .LingH * CantidadItems, Userindex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, .LingP * CantidadItems, Userindex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, .LingO * CantidadItems, Userindex)

    End With

End Sub

Sub CarpinteroQuitarMateriales(ByVal Userindex As Integer, _
                               ByVal ItemIndex As Integer, _
                               ByVal CantidadItems As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora quita tambien madera elfica
    '***************************************************
    With ObjData(ItemIndex)

        If .Madera > 0 Then Call QuitarObjetos(Lena, .Madera * CantidadItems, Userindex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LenaElfica, .MaderaElfica * CantidadItems, Userindex)

    End With

End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, _
                                   ByVal ItemIndex As Integer, _
                                   ByVal Cantidad As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    
    With ObjData(ItemIndex)

        If .Madera > 0 Then
            If Not TieneObjetos(Lena, .Madera * Cantidad, Userindex) Then
                If ShowMsg Then Call WriteConsoleMsg(Userindex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function

            End If

        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LenaElfica, .MaderaElfica * Cantidad, Userindex) Then
                If ShowMsg Then Call WriteConsoleMsg(Userindex, "No tienes suficiente madera elfica.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function

            End If

        End If
    
    End With

    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal Userindex As Integer, _
                                ByVal ItemIndex As Integer, _
                                ByVal CantidadItems As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, .LingH * CantidadItems, Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, .LingP * CantidadItems, Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, .LingO * CantidadItems, Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

    End With

    HerreroTieneMateriales = True

End Function

Function TieneMaterialesUpgrade(ByVal Userindex As Integer, _
                                ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 12/08/2009
    '
    '***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)

        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function

            End If

        End If
        
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function

            End If

        End If
        
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function

            End If

        End If
        
        If .Madera > 0 Then
            If Not TieneObjetos(Lena, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function

            End If

        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LenaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficiente madera elfica.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function

            End If

        End If

    End With
    
    TieneMaterialesUpgrade = True

End Function

Sub QuitarMaterialesUpgrade(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 12/08/2009
    '
    '***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)

        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .Madera > 0 Then Call QuitarObjetos(Lena, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LenaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), Userindex)

    End With
    
    Call QuitarObjetos(ItemIndex, 1, Userindex)

End Sub

Public Function PuedeConstruir(ByVal Userindex As Integer, _
                               ByVal ItemIndex As Integer, _
                               ByVal CantidadItems As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/08/2009
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
    '***************************************************
    PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex, CantidadItems) And Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).Clase), 0) >= ObjData(ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 30/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '30/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    Dim CantidadItems   As Integer

    Dim TieneMateriales As Boolean

    Dim OtroUserIndex   As Integer

    With UserList(Userindex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
                Call LimpiarComercioSeguro(Userindex)

            End If

        End If
        
        CantidadItems = .Construir.PorCiclo
    
        If .Construir.Cantidad < CantidadItems Then CantidadItems = .Construir.Cantidad
        
        If .Construir.Cantidad > 0 Then .Construir.Cantidad = .Construir.Cantidad - CantidadItems
        
        If CantidadItems = 0 Then
            Call WriteStopWorking(Userindex)
            Exit Sub

        End If
    
        If PuedeConstruirHerreria(ItemIndex) Then
        
            While CantidadItems > 0 And Not TieneMateriales

                If PuedeConstruir(Userindex, ItemIndex, CantidadItems) Then
                    TieneMateriales = True
                Else
                    CantidadItems = CantidadItems - 1

                End If

            Wend
        
            ' Chequeo si puede hacer al menos 1 item
            If Not TieneMateriales Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
                Call WriteStopWorking(Userindex)
                Exit Sub

            End If
        
            'Sacamos energia
            If .Clase = eClass.Worker Then

                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                    Call WriteUpdateSta(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            Else

                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                    Call WriteUpdateSta(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
            Call HerreroQuitarMateriales(Userindex, ItemIndex, CantidadItems)
            ' AGREGAR FX
        
            Select Case ObjData(ItemIndex).OBJType
        
                Case eOBJType.otWeapon
                    Call WriteConsoleMsg(Userindex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armas!", "el arma!"), FontTypeNames.FONTTYPE_INFO)

                Case eOBJType.otEscudo
                    Call WriteConsoleMsg(Userindex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " escudos!", "el escudo!"), FontTypeNames.FONTTYPE_INFO)

                Case Is = eOBJType.otCasco
                    Call WriteConsoleMsg(Userindex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " cascos!", "el casco!"), FontTypeNames.FONTTYPE_INFO)

                Case eOBJType.otArmadura
                    Call WriteConsoleMsg(Userindex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armaduras", "la armadura!"), FontTypeNames.FONTTYPE_INFO)
        
            End Select
        
            Dim MiObj As obj
        
            MiObj.Amount = CantidadItems
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
        
            Call SubirSkill(Userindex, eSkill.Herreria, True)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TRABAJO_HERRERO, .Pos.X, .Pos.Y))
        
            If Not criminal(Userindex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
        
            .Counters.Trabajando = .Counters.Trabajando + 1

        End If

    End With

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)

        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function

        End If

    Next i

    PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 28/05/2010
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo errHandler

    Dim CantidadItems   As Integer

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(Userindex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(Userindex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> SERRUCHO_CARPINTERO And WeaponIndex <> SERRUCHO_CARPINTERO_NEWBIE Then
            Call WriteConsoleMsg(Userindex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(Userindex)
            Exit Sub

        End If
        
        CantidadItems = .Construir.PorCiclo
        
        If .Construir.Cantidad < CantidadItems Then CantidadItems = .Construir.Cantidad
            
        If .Construir.Cantidad > 0 Then .Construir.Cantidad = .Construir.Cantidad - CantidadItems
            
        If CantidadItems = 0 Then
            Call WriteStopWorking(Userindex)
            Exit Sub

        End If
    
        If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.Clase), 0) >= ObjData(ItemIndex).SkCarpinteria And PuedeConstruirCarpintero(ItemIndex) Then
           
            ' Calculo cuantos item puede construir
            While CantidadItems > 0 And Not TieneMateriales

                If CarpinteroTieneMateriales(Userindex, ItemIndex, CantidadItems) Then
                    TieneMateriales = True
                Else
                    CantidadItems = CantidadItems - 1

                End If

            Wend
            
            ' No tiene los materiales ni para construir 1 item?
            If Not TieneMateriales Then
                ' Para que muestre el mensaje
                Call CarpinteroTieneMateriales(Userindex, ItemIndex, 1, True)
                Call WriteStopWorking(Userindex)
                Exit Sub

            End If
           
            'Sacamos energia
            If .Clase = eClass.Worker Then

                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                    Call WriteUpdateSta(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            Else

                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                    Call WriteUpdateSta(Userindex)
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
            
            Call CarpinteroQuitarMateriales(Userindex, ItemIndex, CantidadItems)
            Call WriteConsoleMsg(Userindex, "Has construido " & CantidadItems & IIf(CantidadItems = 1, " objeto!", " objetos!"), FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = CantidadItems
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
            
            Call SubirSkill(Userindex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(Userindex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en CarpinteroConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & Userindex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub ArtesanoConstruirItem(ByVal Userindex As Integer, ByVal Item As Integer)
    Dim ArtesanoObj As ObjData
    ArtesanoObj = ObjData(ObjArtesano(Item))

    Dim NpcIndex As Integer
    NpcIndex = UserList(Userindex).flags.TargetNPC

    ' Revisamos si tiene las monedas para la comision
    If UserList(Userindex).Stats.Gld < ArtesaniaCosto Then
        Call WriteChatOverHead(Userindex, "No tienes suficientes monedas de oro para pagarme!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        Exit Sub
    End If

    ' Revisamos si tiene los materiales necesarios
    Dim i As Integer
    For i = 1 To UBound(ArtesanoObj.ItemCrafteo)

        With ArtesanoObj.ItemCrafteo(i)

            If Not TieneObjetos(.ObjIndex, .Amount, Userindex) Then
                Call WriteChatOverHead(Userindex, "No tienes los materiales necesarios!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                Exit Sub
            End If

        End With

    Next i

    ' Le sacamos el oro
    UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - ArtesaniaCosto
    Call WriteUpdateGold(Userindex)
    
    Call WriteConsoleMsg(Userindex, "Le has pagado " & Format$(ArtesaniaCosto, "##,##") & " monedas de oro al artesano.", FontTypeNames.FONTTYPE_INFO)

    ' Le sacamos los materiales
    For i = 1 To UBound(ArtesanoObj.ItemCrafteo)

        With ArtesanoObj.ItemCrafteo(i)

            Call QuitarObjetos(.ObjIndex, .Amount, Userindex)

        End With

    Next i

    Dim ObjetoCreado As obj
    ObjetoCreado.ObjIndex = ObjArtesano(Item)
    ObjetoCreado.Amount = 1

    ' Metemos el item en el inventario o lo tiramos al piso
    If Not MeterItemEnInventario(Userindex, ObjetoCreado) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, ObjetoCreado)
    End If

    Call WriteChatOverHead(Userindex, "Aqui tienes tu " & ArtesanoObj.Name & ". Vuelve pronto!", Npclist(NpcIndex).Char.CharIndex, vbWhite)

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14

        Case iMinerales.PlataCruda
            MineralesParaLingote = 20

        Case iMinerales.OroCrudo
            MineralesParaLingote = 35

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim Slot           As Integer

    Dim obji           As Integer

    Dim CantidadItems  As Integer

    Dim TieneMinerales As Boolean

    Dim OtroUserIndex  As Integer
    
    With UserList(Userindex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(Userindex)

            End If

        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0

        End If
        
        Dim MiObj As obj

        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
        
        Call UpdateUserInv(False, Userindex, Slot)
        Call WriteConsoleMsg(Userindex, "Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub DoFundir(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2010
    '03/06/2010 - Pato: Si es el ultimo item a fundir y esta equipado lo desequipamos.
    '11/03/2010 - ZaMa: Reemplazo division por producto para uan mejor performanse.
    '***************************************************
    Dim i             As Integer

    Dim Num           As Integer

    Dim Slot          As Byte

    Dim Lingotes(2)   As Integer

    Dim OtroUserIndex As Integer

    With UserList(Userindex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(Userindex)

            End If

        End If
        
        Slot = .flags.TargetObjInvSlot
        
        With .Invent.Object(Slot)
            .Amount = .Amount - 1
            
            If .Amount < 1 Then
                If .Equipped = 1 Then Call Desequipar(Userindex, Slot)
                
                .Amount = 0
                .ObjIndex = 0

            End If

        End With
        
        Num = RandomNumber(10, 25)
        
        Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
        Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
        Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
    
        Dim MiObj(2) As obj
        
        For i = 0 To 2
            MiObj(i).Amount = Lingotes(i)
            MiObj(i).ObjIndex = LingoteHierro + i 'Una gran negrada pero practica
            
            If MiObj(i).Amount > 0 Then
                If Not MeterItemEnInventario(Userindex, MiObj(i)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(i))

                End If

            End If

        Next i
        
        Call UpdateUserInv(False, Userindex, Slot)
        Call WriteConsoleMsg(Userindex, "Has obtenido el " & Num & "% de los lingotes utilizados para la construccion del objeto!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub DoUpgrade(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 12/08/2009
    '12/08/2009: Pato - Implementado nuevo sistema de mejora de items
    '***************************************************
    Dim ItemUpgrade   As Integer

    Dim WeaponIndex   As Integer

    Dim OtroUserIndex As Integer

    ItemUpgrade = ObjData(ItemIndex).Upgrade

    With UserList(Userindex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
                Call LimpiarComercioSeguro(Userindex)

            End If

        End If
        
        'Sacamos energia
        If .Clase = eClass.Worker Then

            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                Call WriteUpdateSta(Userindex)
            Else
                Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        Else

            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                Call WriteUpdateSta(Userindex)
            Else
                Call WriteConsoleMsg(Userindex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
    
        If ItemUpgrade <= 0 Then Exit Sub
        If Not TieneMaterialesUpgrade(Userindex, ItemIndex) Then Exit Sub
    
        If PuedeConstruirHerreria(ItemUpgrade) Then
        
            WeaponIndex = .Invent.WeaponEqpObjIndex
    
            If WeaponIndex <> MARTILLO_HERRERO And WeaponIndex <> MARTILLO_HERRERO_NEWBIE Then
                Call WriteConsoleMsg(Userindex, "Debes equiparte el martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If Round(.Stats.UserSkills(eSkill.Herreria) / ModHerreriA(.Clase), 0) < ObjData(ItemUpgrade).SkHerreria Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Select Case ObjData(ItemIndex).OBJType

                Case eOBJType.otWeapon
                    Call WriteConsoleMsg(Userindex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
                Case eOBJType.otEscudo 'Todavia no hay, pero just in case
                    Call WriteConsoleMsg(Userindex, "Has mejorado el escudo!", FontTypeNames.FONTTYPE_INFO)
            
                Case eOBJType.otCasco
                    Call WriteConsoleMsg(Userindex, "Has mejorado el casco!", FontTypeNames.FONTTYPE_INFO)
            
                Case eOBJType.otArmadura
                    Call WriteConsoleMsg(Userindex, "Has mejorado la armadura!", FontTypeNames.FONTTYPE_INFO)

            End Select
        
            Call SubirSkill(Userindex, eSkill.Herreria, True)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TRABAJO_HERRERO, .Pos.X, .Pos.Y))
    
        ElseIf PuedeConstruirCarpintero(ItemUpgrade) Then
        
            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex <> SERRUCHO_CARPINTERO And WeaponIndex <> SERRUCHO_CARPINTERO_NEWBIE Then
                Call WriteConsoleMsg(Userindex, "Debes equiparte un serrucho.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.Clase), 0) < ObjData(ItemUpgrade).SkCarpinteria Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Select Case ObjData(ItemIndex).OBJType

                Case eOBJType.otFlechas
                    Call WriteConsoleMsg(Userindex, "Has mejorado la flecha!", FontTypeNames.FONTTYPE_INFO)
                
                Case eOBJType.otWeapon
                    Call WriteConsoleMsg(Userindex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
                Case eOBJType.otBarcos
                    Call WriteConsoleMsg(Userindex, "Has mejorado el barco!", FontTypeNames.FONTTYPE_INFO)

            End Select
        
            Call SubirSkill(Userindex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
        Else
            Exit Sub

        End If
    
        Call QuitarMaterialesUpgrade(Userindex, ItemIndex)
    
        Dim MiObj As obj

        MiObj.Amount = 1
        MiObj.ObjIndex = ItemUpgrade
    
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
    
        If ObjData(ItemIndex).Log = 1 Then Call LogDesarrollo(.Name & " ha mejorado el item " & ObjData(ItemIndex).Name & " a " & ObjData(ItemUpgrade).Name)
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Function ModFundicion(ByVal Clase As eClass) As Single
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Clase

        Case eClass.Worker
            ModFundicion = 1

        Case Else
            ModFundicion = 3

    End Select

End Function

Function ModCarpinteria(ByVal Clase As eClass) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Clase

        Case eClass.Worker
            ModCarpinteria = 1

        Case Else
            ModCarpinteria = 3

    End Select

End Function

Function ModHerreriA(ByVal Clase As eClass) As Single

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Clase

        Case eClass.Worker
            ModHerreriA = 1

        Case Else
            ModHerreriA = 4

    End Select

End Function

Function ModDomar(ByVal Clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Clase

        Case eClass.Druid
            ModDomar = 6

        Case eClass.Hunter
            ModDomar = 6

        Case eClass.Cleric
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(Userindex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo errHandler

    Dim puntosDomar      As Integer

    Dim puntosRequeridos As Integer

    Dim CanStay          As Boolean

    Dim petType          As Integer

    Dim NroPets          As Integer
    
    If Npclist(NpcIndex).MaestroUser = Userindex Then
        Call WriteConsoleMsg(Userindex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    With UserList(Userindex)

        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(Userindex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If Not PuedeDomarMascota(Userindex, NpcIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            
            ' 20% de bonificacion
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
            
                ' 11% de bonificacion
            ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.89
                
            Else
                puntosRequeridos = Npclist(NpcIndex).flags.Domable

            End If
            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                Dim index As Integer

                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(Userindex)
                .MascotasIndex(index) = NpcIndex
                .MascotasType(index) = Npclist(NpcIndex).Numero
                
                Npclist(NpcIndex).MaestroUser = Userindex
                
                Call FollowAmo(NpcIndex)
                Call ReSpawnNpc(Npclist(NpcIndex))
                
                Call WriteConsoleMsg(Userindex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
                
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.Map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(NpcIndex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(NpcIndex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                    Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call SubirSkill(Userindex, eSkill.Domar, True)
        
            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(Userindex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5

                End If
                
                Call SubirSkill(Userindex, eSkill.Domar, False)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal Userindex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long

    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(Userindex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    
    Dim tempData As String
    
    With UserList(Userindex)

        If .flags.AdminInvisible = 0 Then

            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False

            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            
            ' Solo el admin sabe que se hace invi
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, True)
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(tempData)
            
            'Le mandamos el mensaje para que borre el personaje a los clientes que esten cerca
            Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
            
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            ' Solo el admin sabe que se hace visible
            tempData = PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim)
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(tempData)
            
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, False)
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(tempData)
             
            'Le mandamos el mensaje para crear el personaje a los clientes que esten cerca
            Call MakeUserChar(True, .Pos.Map, Userindex, .Pos.Map, .Pos.X, .Pos.Y, True)

        End If

    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Suerte    As Byte

    Dim exito     As Byte

    Dim obj       As obj

    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteConsoleMsg(Userindex, "Necesitas clickear sobre lena para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Distancia(posMadera, UserList(Userindex).Pos) > 2 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(Userindex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(Userindex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    Dim SupervivenciaSkill As Byte

    SupervivenciaSkill = UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia)

    If SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill <= 34 Then
        Suerte = 2
    Else
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(Userindex, "Has hecho " & obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(Userindex).flags.TargetObj = FOGATA_APAG
    
        Call SubirSkill(Userindex, eSkill.Supervivencia, True)
    Else

        '[CDT 17-02-2004]
        If Not UserList(Userindex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(Userindex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    
        Call SubirSkill(Userindex, eSkill.Supervivencia, False)

    End If

End Sub

Public Sub DoPescar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/10/2018
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim Skill         As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(Userindex)

        If .Clase = eClass.Worker Then
            Call QuitarSta(Userindex, EsfuerzoPescarPescador)
        Else
            Call QuitarSta(Userindex, EsfuerzoPescarGeneral)

        End If
    
        Skill = .Stats.UserSkills(eSkill.pesca)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= DificultadPescar Then

            Dim MiObj As obj
        
            If .Clase = eClass.Worker Then
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
                CantidadItems = RandomNumber(1, MAXITEMS)
            Else
                CantidadItems = 1

            End If

            CantidadItems = CantidadItems * OficioMultiplier

            Dim i As Long

            With MiObj
    
                If PescaEvent.Activado = 1 Then

                    For i = 1 To PescaEvent.CantidadDeZonas

                        If UserList(Userindex).Pos.Map = Zona(i).Mapa Then
                            MiObj.ObjIndex = Evento_Pesca.DamePez(i)
                        Else
                            .Amount = CantidadItems
                            MiObj.ObjIndex = Pescado
                        End If

                    Next i

                Else
                    .Amount = CantidadItems
                    MiObj.ObjIndex = Pescado
                End If

            End With

            'With MiObj
            '    .Amount = CantidadItems
            '    .ObjIndex = Pescado
            'End With
            
            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            Call WriteConsoleMsg(Userindex, "Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            
            Call SubirSkill(Userindex, eSkill.pesca, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 6 Then
                Call WriteConsoleMsg(Userindex, "No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 6

            End If

            '[/CDT]
        
            Call SubirSkill(Userindex, eSkill.pesca, False)

        End If
    
        If Not criminal(Userindex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.description)

End Sub

Public Sub DoPescarRed(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/05/2020
    '26/10/2018: CHOTS - Multiplicador de oficios
    '26/05/2020: Fakkerz - Agregado render de dano faltante
    '***************************************************
    On Error GoTo errHandler

    Dim iSkill        As Integer

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim EsPescador    As Boolean

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(Userindex)
    
        If .Clase = eClass.Worker Then
            Call QuitarSta(Userindex, EsfuerzoPescarPescador)
            EsPescador = True
        Else
            Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
            EsPescador = False

        End If
        
        iSkill = .Stats.UserSkills(eSkill.pesca)
        
        ' m = (60-11)/(1-10)
        ' y = mx - m*10 + 11
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= DificultadPescar Then
            
                Dim MiObj As obj
                
                If EsPescador Then
                    MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
                    CantidadItems = RandomNumber(1, MAXITEMS)
                Else
                    CantidadItems = 1

                End If

                CantidadItems = CantidadItems * OficioMultiplier
                
                MiObj.Amount = CantidadItems
                MiObj.ObjIndex = ListaPeces(RandomNumber(1, NUM_PECES))
                
                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call WriteConsoleMsg(Userindex, "Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
                
                'Renderizo el dano en render.
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
                
                Call SubirSkill(Userindex, eSkill.pesca, True)
            Else

                If Not .flags.UltimoMensaje = 6 Then
                    Call WriteConsoleMsg(Userindex, "No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 6

                End If
                
                Call SubirSkill(Userindex, eSkill.pesca, False)

            End If

        End If
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    
    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoPescarRed")

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************

    On Error GoTo errHandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

        Else

            If .Faccion.ArmadaReal = 1 Then
                If Not criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            End If

        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer

            Dim res        As Integer

            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)

                    End If

                End If
               
                If (RandomNumber(1, 50) < 25) And (.Clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.Gld > 0 Then

                        Dim n As Long
                        
                        If .Clase = eClass.Thief Then

                            ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                n = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)

                            End If

                        Else
                            n = RandomNumber(1, 100)

                        End If

                        If n > UserList(VictimaIndex).Stats.Gld Then n = UserList(VictimaIndex).Stats.Gld
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - n
                        
                        .Stats.Gld = .Stats.Gld + n

                        If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

            End If
        
            If Not criminal(LadrOnIndex) Then
                If Not criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)

                End If

            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

                If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

            End If

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregue los barcos
    ' Esta funcion determina que objetos son robables.
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And Not ItemNewbie(OI)

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************

    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    With UserList(VictimaIndex)

        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
            i = 1

            Do While Not flag And i <= .CurrentInventorySlots

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i + 1
            Loop
        Else
            i = .CurrentInventorySlots

            Do While Not flag And i > 0

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i - 1
            Loop

        End If
    
        If flag Then

            Dim MiObj     As obj

            Dim Num       As Integer

            Dim ObjAmount As Integer
        
            ObjAmount = .Invent.Object(i).Amount
        
            'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
            Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
            MiObj.Amount = Num
            MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
            .Invent.Object(i).Amount = ObjAmount - Num
                    
            If .Invent.Object(i).Amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

            End If
                
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
            If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
        
            If UserList(LadrOnIndex).Clase = eClass.Thief Then
                Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)

    End With

End Sub

Public Sub DoApunalar(ByVal Userindex As Integer, _
                      ByVal VictimNpcIndex As Integer, _
                      ByVal VictimUserIndex As Integer, _
                      ByVal dano As Long)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el dano
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apunalar)

    Select Case UserList(Userindex).Clase

        Case eClass.Assasin
            Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Cleric, eClass.Paladin, eClass.Pirat
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bard
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(Userindex).Clase = eClass.Assasin Then
                dano = Round(dano * 1.4, 0)
            Else
                dano = Round(dano * 1.5, 0)

            End If
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, dano, DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(Userindex, "Has apunalado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, "Te ha apunalado " & UserList(Userindex).Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
        
        Else
            
            With Npclist(VictimNpcIndex)
                .Stats.MinHp = .Stats.MinHp - Int(dano * 2)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(Userindex, "Has apunalado la criatura por " & Int(dano * 2), FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(Userindex, VictimNpcIndex, dano * 2)
            
            End With

        End If
    
        Call SubirSkill(Userindex, eSkill.Apunalar, True)
    Else
        Call WriteConsoleMsg(Userindex, "No has logrado apunalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(Userindex, eSkill.Apunalar, False)

    End If

End Sub

Public Sub DoAcuchillar(ByVal Userindex As Integer, _
                        ByVal VictimNpcIndex As Integer, _
                        ByVal VictimUserIndex As Integer, _
                        ByVal dano As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 12/01/2010
    '***************************************************

    If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
        dano = Int(dano * DANO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                Call WriteConsoleMsg(Userindex, "Has acuchillado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).Name & " te ha acuchillado por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
        
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - dano
            Call WriteConsoleMsg(Userindex, "Has acuchillado a la criatura por " & dano, FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(Userindex, VictimNpcIndex, dano)
        
        End If

    End If
    
End Sub

Public Sub DoGolpeCritico(ByVal Userindex As Integer, _
                          ByVal VictimNpcIndex As Integer, _
                          ByVal VictimUserIndex As Integer, _
                          ByVal dano As Long)

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
    '***************************************************
    Dim Suerte      As Integer

    Dim Skill       As Integer

    Dim WeaponIndex As Integer
    
    With UserList(Userindex)

        ' Es bandido?
        If .Clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Wrestling)

    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        dano = Int(dano * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(Userindex, "Has golpeado criticamente a " & .Name & " por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).Name & " te ha golpeado criticamente por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            
            With Npclist(VictimNpcIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(Userindex, "Has golpeado criticamente a la criatura por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                
                Call CalcularDarExp(Userindex, VictimNpcIndex, dano)
            End With
            
           
            
        End If
        
    End If

End Sub

Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Cantidad

    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call WriteUpdateSta(Userindex)
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal Userindex As Integer, _
                   Optional ByVal DarMaderaElfica As Boolean = False)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 26/10/2018
    '16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    Dim Skill         As Integer

    With UserList(Userindex)

        If .Clase = eClass.Worker Then
            Call QuitarSta(Userindex, EsfuerzoTalarLenador)
        Else
            Call QuitarSta(Userindex, EsfuerzoTalarGeneral)

        End If
    
        Skill = .Stats.UserSkills(eSkill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= DificultadTalar Then

            Dim MiObj As obj
        
            If .Clase = eClass.Worker Then
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
                CantidadItems = RandomNumber(1, MAXITEMS)
            Else
                CantidadItems = 1

            End If

            CantidadItems = CantidadItems * OficioMultiplier
            
            With MiObj
                .Amount = CantidadItems
                .ObjIndex = IIf(DarMaderaElfica, LenaElfica, Lena)
            End With
            
        
            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            Call WriteConsoleMsg(Userindex, "Has conseguido algo de lena!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            
            Call SubirSkill(Userindex, eSkill.Talar, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 8 Then
                Call WriteConsoleMsg(Userindex, "No has obtenido lena!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 8

            End If

            '[/CDT]
            Call SubirSkill(Userindex, eSkill.Talar, False)

        End If
    
        If Not criminal(Userindex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal Userindex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 26/10/2018
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(Userindex)

        If .Clase = eClass.Worker Then
            Call QuitarSta(Userindex, EsfuerzoExcavarMinero)
        Else
            Call QuitarSta(Userindex, EsfuerzoExcavarGeneral)

        End If
    
        Dim Skill As Integer

        Skill = .Stats.UserSkills(eSkill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= DificultadMinar Then

            Dim MiObj As obj
        
            If .flags.TargetObj = 0 Then Exit Sub
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
        
            If .Clase = eClass.Worker Then
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
                CantidadItems = RandomNumber(1, MAXITEMS)
            Else
                CantidadItems = 1

            End If

            CantidadItems = CantidadItems * OficioMultiplier

            MiObj.Amount = CantidadItems
        
            If Not MeterItemEnInventario(Userindex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(Userindex, "Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            
            Call SubirSkill(Userindex, eSkill.Mineria, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(Userindex, "No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(Userindex, eSkill.Mineria, False)

        End If
    
        If Not criminal(Userindex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

errHandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(Userindex)
        .Counters.IdleCount = 0
        
        Dim Suerte       As Integer

        Dim res          As Integer

        Dim cant         As Integer

        Dim MeditarSkill As Byte
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual      As Long

        TActual = GetTickCount()

        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub

        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True

        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(Userindex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(Userindex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Exit Sub

        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 Then
            Suerte = 7
        Else
            Suerte = 5

        End If
        
        ' WyroX: Si, un poco rudimentario, pero deberia ayudar a equilibrar la cantidad de mana "perdida"
        ' cuando los ticks del juego no son exactamente de 40ms
        res = RandomNumber(1, Round(Suerte / DeltaTick))
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
            If Not .flags.UltimoMensaje = 22 Then
                Call WriteConsoleMsg(Userindex, "Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 22

            End If
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, cant, DAMAGE_TRABAJO))
            
            Call WriteUpdateMana(Userindex)
            Call SubirSkill(Userindex, eSkill.Meditar, True)
        Else
            Call SubirSkill(Userindex, eSkill.Meditar, False)

        End If

    End With

End Sub

Public Sub DoDesequipar(ByVal Userindex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte

    Dim AlgoEquipado   As Boolean
    
    With UserList(Userindex)

        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66

    End With
   
    With UserList(VictimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(Userindex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
    '***************************************************
    Dim OtroUserIndex As Integer

    If TriggerZonaPelea(Userindex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(Userindex).Clase <> eClass.Bandit Then Exit Sub

    'Esto es precario y feo, pero por ahora no se me ocurrio nada mejor.
    'Uso el slot de los anillos para "equipar" los guantes.
    'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
    If UserList(Userindex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

    Dim res As Integer

    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
    
            If UserList(VictimaIndex).flags.Comerciando Then
                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                    Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
                    Call LimpiarComercioSeguro(VictimaIndex)

                End If

            End If
                
            Call RobarObjeto(Userindex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "" & UserList(Userindex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub DoHandInmo(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 17/02/2007
    'Implements the special Skill of the Thief
    '***************************************************
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    If UserList(Userindex).Clase <> eClass.Thief Then Exit Sub
    
    If UserList(Userindex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
    Dim res As Integer

    res = RandomNumber(0, 100)

    If res < (UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
        UserList(VictimaIndex).flags.ParalizedByIndex = Userindex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(Userindex).Name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(Userindex, "Tu golpe ha dejado inmovil a tu oponente", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimaIndex, "El golpe te ha dejado inmovil!", FontTypeNames.FONTTYPE_FIGHT)

    End If

End Sub

Public Sub Desarmar(ByVal Userindex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte
    
    With UserList(Userindex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal Userindex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '***************************************************
    
    With UserList(Userindex)

        If .Clase = eClass.Worker Then
            MaxItemsConstruibles = MaximoInt(1, CInt((.Stats.ELV - 2) * 0.2))
        Else
            MaxItemsConstruibles = 1

        End If

    End With

End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/05/2010
    '***************************************************
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1

End Function

Public Sub ImitateNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Copies body, head and desc from previously clicked npc.
    '***************************************************
    
    With UserList(Userindex)
        
        ' Copy desc
        .DescRM = Npclist(NpcIndex).Name
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NpcIndex).Char.body
            .flags.OldHead = Npclist(NpcIndex).Char.Head
        Else
            .Char.body = Npclist(NpcIndex).Char.body
            .Char.Head = Npclist(NpcIndex).Char.Head
            
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End If
    
    End With
    
End Sub

Public Sub DoEquita(ByVal Userindex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)
    '***************************************************
    'Author: Recox
    'Last Modification: 06/04/2020
    'Podemos usar monturas ahora
    '06/04/2020: FrankoH298 - Ahora hay un timer para poder montarte
    '***************************************************
    'Dim ModEqui As Long

    'ModEqui = ModEquitacion(UserList(UserIndex).clase)

    ' Comento esto por que aun no implementamos en el frmSkills la posibilidad de agregar skills de equitacion
    'If UserList(UserIndex).Stats.UserSkills(Equitacion) / ModEqui < Montura.MinSkill Then
    '    Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill * ModEqui & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
    '    Exit Sub
    'End If

    With UserList(Userindex)

        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes utilizar la montura mientras estas muerto !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes utilizar la montura mientras navegas !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapInfo(.Pos.Map).Zona = Dungeon Or _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.BAJOTECHO Or _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.CASA Then
           
            If .flags.Equitando = 0 Then
                Call WriteConsoleMsg(Userindex, "No puedes utilizar la montura bajo techo o dungeons!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
        End If

        ' If .flags.Metamorfosis = 1 Then 'Metamorfosis
        '     Call WriteConsoleMsg(UserIndex, "No puedes montar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
        '     Exit Sub
        ' End If

        ' No estaba equitando
        If .flags.Equitando = 0 Then

            If .Counters.MonturaCounter <= 0 Then
                .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.MonturaEqpSlot = Slot
    
                Call ToggleMonturaBody(Userindex)
                Call SetVisibleStateForUserAfterNavigateOrEquitate(Userindex)
    
                '  Comienza a equitar
                .flags.Equitando = 1

                Call WriteEquitandoToggle(Userindex)

                'Mostramos solo el casco de los items equipados por que los demas items quedan mal en el render, solo es un tema visual (Recox)
                Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, .Char.CascoAnim)
            Else
                Call WriteConsoleMsg(Userindex, "Debe esperar " & .Counters.MonturaCounter & " segundos para volver a usar tu montura", FontTypeNames.FONTTYPE_INFO)
            End If
            
        ' Estaba equitando
        Else
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)

        End If


    End With

End Sub

Public Sub UnmountMontura(ByVal Userindex As Integer)
    With UserList(Userindex)
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0

        .Char.Head = .OrigChar.Head

        ' Seteamos el equipo que tiene y lo mostramos en el render.
        Call SetEquipmentOnCharAfterNavigateOrEquitate(Userindex)
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
  
        ' Termina de equitar
        .flags.Equitando = 0
        .Counters.MonturaCounter = 10

    End With
End Sub

Private Function ModEquitacion(ByVal UserClase As Byte) As Integer

    Select Case UserClase
        Case eClass.Cleric
            ModEquitacion = 1
        Case Else
            ModEquitacion = 1.5
    End Select

End Function

Private Sub SetVisibleStateForUserAfterNavigateOrEquitate(ByVal Userindex As Integer)

    With UserList(Userindex)

        ' Pierde el ocultar
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.Ocultando = 0
            Call SetInvisible(Userindex, .Char.CharIndex, False)
            Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Siempre se ve la montura (Nunca esta invisible), pero solo para el cliente.
        If .flags.invisible = 1 Then
            Call SetInvisible(Userindex, .Char.CharIndex, False)
        End If

    End With

End Sub

Private Sub SetEquipmentOnCharAfterNavigateOrEquitate(ByVal Userindex As Integer)

    With UserList(Userindex)
        
        If .flags.Mimetizado = 0 Then
            If .Invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(Userindex)
            End If
        End If
        
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

        If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(Userindex, .Invent.WeaponEqpObjIndex)

        If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
        
    End With


End Sub
