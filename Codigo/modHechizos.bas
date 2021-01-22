Attribute VB_Name = "modHechizos"
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

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, _
                           ByVal Userindex As Integer, _
                           ByVal Spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
    '13/07/2010: ZaMa - Ahora no se contabiliza la muerte de un atacable.
    '21/09/2010: ZaMa - Amplio los tipos de hechizos que pueden lanzar los npcs.
    '21/09/2010: ZaMa - Permito que se ignore el chequeo de visibilidad (pueden atacar a invis u ocultos).
    '11/11/2010: ZaMa - No se envian los efectos del hechizo si no lo castea.
    '06/04/2020: FrankoH298 - Si te lanzan un hechizo te desmonta
    '***************************************************

    If Not IntervaloPermiteAtacarNpc(NpcIndex) Then Exit Sub

    With UserList(Userindex)
    
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
            
        End If
        
        ' Doesn't consider if the user is hidden/invisible or not.
        If Not IgnoreVisibilityCheck Then
            If UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then Exit Sub

        End If
        
        ' Si no se peude usar magia en el mapa, no le deja hacerlo.
        If MapInfo(UserList(Userindex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

        Dim dano As Integer
    
        ' Heal HP
        If Hechizos(Spell).SubeHP = 1 Then
        
            Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
        
            dano = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
            .Stats.MinHp = .Stats.MinHp + dano

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
            
            Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            
            Call WriteUpdateUserStats(Userindex)
        
            ' Damage
        ElseIf Hechizos(Spell).SubeHP = 2 Then
            
            If .flags.Privilegios And PlayerType.User Then
            
                Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
            
                dano = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
                
                If .Invent.CascoEqpObjIndex > 0 Then
                    dano = dano - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

                End If
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    dano = dano - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

                End If
                
                If dano < 0 Then dano = 0
            
                .Stats.MinHp = .Stats.MinHp - dano
                
                Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render.
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
                
                Call WriteUpdateUserStats(Userindex)
                
                'Muere
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0

                    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                        RestarCriminalidad (Userindex)

                    End If
                    
                    Dim MasterIndex As Integer

                    MasterIndex = Npclist(NpcIndex).MaestroUser
                    
                    '[Barrin 1-12-03]
                    If MasterIndex > 0 Then
                        
                        ' No son frags los muertos atacables
                        If .flags.AtacablePor <> MasterIndex Then
                            'Store it!
                            Call Statistics.StoreFrag(MasterIndex, Userindex)
                            
                            Call ContarMuerte(Userindex, MasterIndex)

                        End If
                        
                        Call ActStats(Userindex, MasterIndex)

                    End If

                    '[/Barrin]
                    
                    Call UserDie(Userindex)
                    
                End If
            
            End If
            
        End If
        
        ' Paralisis/Inmobilize
        If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
        
            If .flags.Paralizado = 0 Then
                
                Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideParalizar Then
                        Call WriteConsoleMsg(Userindex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                
                If Hechizos(Spell).Inmoviliza = 1 Then
                    .flags.Inmovilizado = 1

                End If
                  
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
                  
                Call WriteParalizeOK(Userindex)
                
            End If
            
        End If
        
        ' Stupidity
        If Hechizos(Spell).Estupidez = 1 Then
             
            If .flags.Estupidez = 0 Then
            
                Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideAturdir Then
                        Call WriteConsoleMsg(Userindex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                  
                .flags.Estupidez = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteDumb(Userindex)
                
            End If

        End If
        
        ' Blind
        If Hechizos(Spell).Ceguera = 1 Then
             
            If .flags.Ceguera = 0 Then
            
                Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideCegar Then
                        Call WriteConsoleMsg(Userindex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                  
                .flags.Ceguera = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteBlind(Userindex)
                
            End If

        End If
        
        ' Remove Invisibility/Hidden
        If Hechizos(Spell).RemueveInvisibilidadParcial = 1 Then
                 
            Call SendSpellEffects(Userindex, NpcIndex, Spell, DecirPalabras)
                 
            'Sacamos el efecto de ocultarse
            If .flags.Oculto = 1 Then
                .Counters.TiempoOculto = 0
                .flags.Oculto = 0
                Call SetInvisible(Userindex, .Char.CharIndex, False)
                Call WriteConsoleMsg(Userindex, "Has sido detectado!", FontTypeNames.FONTTYPE_VENENO)
            Else
                'sino, solo lo "iniciamos" en la sacada de invisibilidad.
                Call WriteConsoleMsg(Userindex, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO)
                .Counters.Invisibilidad = IntervaloInvisible - 1

            End If
        
        End If
        
    End With
    
End Sub

Private Sub SendSpellEffects(ByVal Userindex As Integer, _
                             ByVal NpcIndex As Integer, _
                             ByVal Spell As Integer, _
                             ByVal DecirPalabras As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/12/2016
    'Sends spell's wav, fx and mgic words to users.
    ' Shak: Palabras magicas
    '***************************************************
    With UserList(Userindex)
        ' Spell Wav
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y))
            
        ' Spell FX
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        ' Spell Words
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(Spell, Npclist(NpcIndex).Char.CharIndex))

        End If

    End With

End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, _
                                 ByVal TargetNPC As Integer, _
                                 ByVal SpellIndex As Integer, _
                                 Optional ByVal DecirPalabras As Boolean = False)
    '***************************************************
    'Author: Unknown
    'Last Modification: 21/09/2010
    '21/09/2010: ZaMa - Now npcs can cast a wider range of spells.
    '***************************************************

    If Not IntervaloPermiteAtacarNpc(NpcIndex) Then Exit Sub
    
    Dim Danio As Integer
    
    With Npclist(TargetNPC)
    
        ' Spell sound and FX
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y))
            
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
    
        ' Decir las palabras magicas?
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(SpellIndex, Npclist(NpcIndex).Char.CharIndex))

        End If
    
        ' Spell deals damage??
        If Hechizos(SpellIndex).SubeHP = 2 Then
            
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
            ' Deal damage
            .Stats.MinHp = .Stats.MinHp - Danio
            
            'Muere?
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0

                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNPC, 0)

                End If

            End If
            
            ' Spell recovers health??
        ElseIf Hechizos(SpellIndex).SubeHP = 1 Then
            
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
            ' Recovers health
            .Stats.MinHp = .Stats.MinHp + Danio
            
            If .Stats.MinHp > .Stats.MaxHp Then
                .Stats.MinHp = .Stats.MaxHp

            End If
            
        End If
        
        ' Spell Adds/Removes poison?
        If Hechizos(SpellIndex).Envenena = 1 Then
            .flags.Envenenado = 1
        ElseIf Hechizos(SpellIndex).CuraVeneno = 1 Then
            .flags.Envenenado = 0

        End If

        ' Spells Adds/Removes Paralisis/Inmobility?
        If Hechizos(SpellIndex).Paraliza = 1 Then
            .flags.Paralizado = 1
            .flags.Inmovilizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(SpellIndex).Inmoviliza = 1 Then
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then

            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                .flags.Paralizado = 0
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = 0

            End If

        End If
    
    End With

End Sub

Function TieneHechizo(ByVal i As Integer, ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(Userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
ErrHandler:

End Function

Sub AgregarHechizo(ByVal Userindex As Integer, ByVal Slot As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim hIndex As Integer

    Dim j      As Integer

    With UserList(Userindex)
        hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex
    
        If Not TieneHechizo(hIndex, Userindex) Then

            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS

                If .Stats.UserHechizos(j) = 0 Then Exit For
            Next j
            
            If .Stats.UserHechizos(j) <> 0 Then
                Call WriteConsoleMsg(Userindex, "No tienes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Stats.UserHechizos(j) = hIndex
                Call UpdateUserHechizos(False, Userindex, CByte(j))
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, CByte(Slot), 1)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellIndex As Integer, ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 17/11/2009
    '25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
    '17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
    '28/12/2016: Shak - Palabras magicas
    '21/02/2019: Jopi - Amuleto del Silencio
    '***************************************************
    On Error GoTo ErrHandler
    
    ' Amuleto del Silencio
    If TieneObjetos(AMULETO_DEL_SILENCIO, 1, Userindex) Then Exit Sub
              
    With UserList(Userindex)

        If .flags.AdminInvisible <> 1 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePalabrasMagicas(SpellIndex, .Char.CharIndex))
                
            ' Si estaba oculto, se vuelve visible
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(Userindex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en DecirPalabrasMagicas. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal Userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    'Last Modification By: ZaMa
    '06/11/09 - Corregida la bonificacion de mana del mimetismo en el druida con flauta magica equipada.
    '19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
    '12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
    '***************************************************
    Dim DruidManaBonus As Single

    With UserList(Userindex)

        If .flags.Muerto Then
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Function

        End If
            
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If .Clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(Userindex, "No posees un baculo lo suficientemente poderoso para poder lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "No puedes lanzar este conjuro sin la ayuda de un baculo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If

        End If
            
        If .Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(Userindex, "Estas muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Estas muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Function

        End If
    
        DruidManaBonus = 1

        If .Clase = eClass.Druid Then
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then

                ' 50% menos de mana requerido para mimetismo
                If Hechizos(HechizoIndex).Mimetiza = 1 Then
                    DruidManaBonus = 0.5
                    
                    ' 30% menos de mana requerido para invocaciones
                ElseIf Hechizos(HechizoIndex).Tipo = uInvocacion Then
                    DruidManaBonus = 0.7
                
                    ' 10% menos de mana requerido para las demas magias, excepto apoca
                ElseIf HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                    DruidManaBonus = 0.9

                End If

            End If
            
            ' Necesita tener la barra de mana completa para invocar una mascota
            If Hechizos(HechizoIndex).Warp = 1 Then
                If .Stats.MinMAN <> .Stats.MaxMAN Then
                    Call WriteConsoleMsg(Userindex, "Debes poseer toda tu mana para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                    ' Si no tiene mascotas, no tiene sentido que lo use
                ElseIf .NroMascotas = 0 Then
                    Call WriteConsoleMsg(Userindex, "Debes poseer alguna mascota para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If

        End If
        
        If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
            Call WriteConsoleMsg(Userindex, "No tienes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
    End With
    
    PuedeLanzar = True

End Function

Sub HechizoTerrenoEstado(ByVal Userindex As Integer, ByRef b As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PosCasteadaX As Integer

    Dim PosCasteadaY As Integer

    Dim PosCasteadaM As Integer

    Dim h            As Integer

    Dim tempX        As Integer

    Dim tempY        As Integer

    With UserList(Userindex)
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        h = .flags.Hechizo
        
        If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
            b = True

            For tempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For tempY = PosCasteadaY - 8 To PosCasteadaY + 8

                    If InMapBounds(PosCasteadaM, tempX, tempY) Then
                        If MapData(PosCasteadaM, tempX, tempY).Userindex > 0 Then

                            'hay un user
                            If UserList(MapData(PosCasteadaM, tempX, tempY).Userindex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, tempX, tempY).Userindex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, tempX, tempY).Userindex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                            End If

                        End If

                    End If

                Next tempY
            Next tempX
        
            Call InfoHechizo(Userindex)

        End If

    End With

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operacion.

Sub HechizoInvocacion(ByVal Userindex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Author: Uknown
    'Last modification: 18/09/2010
    'Sale del sub si no hay una posicion valida.
    '18/11/2009: Optimizacion de codigo.
    '18/09/2010: ZaMa - No se permite invocar en mapas con InvocarSinEfecto.
    '***************************************************

    On Error GoTo Error

    With UserList(Userindex)

        Dim Mapa As Integer

        Mapa = .Pos.Map
    
        'No permitimos se invoquen criaturas en zonas seguras
        If MapInfo(Mapa).Pk = False Or MapData(Mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
            Call WriteConsoleMsg(Userindex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'No permitimos se invoquen criaturas en mapas donde esta prohibido hacerlo
        If MapInfo(Mapa).InvocarSinEfecto = 1 Then
            Call WriteConsoleMsg(Userindex, "Invocar no esta permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer

        Dim TargetPos  As WorldPos
    
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
    
        SpellIndex = .flags.Hechizo
    
        ' Warp de mascotas
        If Hechizos(SpellIndex).Warp = 1 Then
            PetIndex = FarthestPet(Userindex)
        
            ' La invoco cerca mio
            If PetIndex > 0 Then
                Call WarpMascota(Userindex, PetIndex)

            End If
        
            ' Invocacion normal
        Else

            If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
            For NroNpcs = 1 To Hechizos(SpellIndex).cant
            
                If .NroMascotas < MAXMASCOTAS Then
                    NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False)

                    If NpcIndex > 0 Then
                        .NroMascotas = .NroMascotas + 1
                    
                        PetIndex = FreeMascotaIndex(Userindex)
                    
                        .MascotasIndex(PetIndex) = NpcIndex
                        .MascotasType(PetIndex) = Npclist(NpcIndex).Numero
                    
                        With Npclist(NpcIndex)
                            .MaestroUser = Userindex
                            .Contadores.TiempoExistencia = IntervaloInvocacion
                            .GiveGLD = 0

                        End With
                    
                        Call FollowAmo(NpcIndex)
                    Else
                        Exit Sub

                    End If

                Else
                    Exit For

                End If
        
            Next NroNpcs

        End If

    End With

    Call InfoHechizo(Userindex)
    HechizoCasteado = True

    Exit Sub

Error:

    With UserList(Userindex)
        LogError ("[" & Err.Number & "] " & Err.description & " por el usuario " & .Name & "(" & Userindex & ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & SpellIndex & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")

    End With

End Sub

Sub HandleHechizoTerreno(ByVal Userindex As Integer, ByVal SpellIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(SpellIndex).Tipo

        Case TipoHechizo.uInvocacion
            Call HechizoInvocacion(Userindex, HechizoCasteado)
            
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(Userindex, HechizoCasteado)

    End Select

    If HechizoCasteado Then

        With UserList(Userindex)
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            If Hechizos(SpellIndex).Warp = 1 Then ' Invoco una mascota
                ' Consume toda la mana
                ManaRequerida = .Stats.MinMAN
            Else

                ' Bonificaciones en hechizos
                If .Clase = eClass.Druid Then

                    ' Solo con flauta equipada
                    If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                        ' 30% menos de mana para invocaciones
                        ManaRequerida = ManaRequerida * 0.7

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)

        End With

    End If
    
End Sub

Sub HandleHechizoUsuario(ByVal Userindex As Integer, ByVal SpellIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
    '***************************************************
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer

    With UserList(Userindex)
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
            
        End If
    End With
    
    Select Case Hechizos(SpellIndex).Tipo

        Case TipoHechizo.uEstado
            ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(Userindex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades
            ' Afectan HP,MANA,STAMINA,ETC
            HechizoCasteado = HechizoPropUsuario(Userindex)

    End Select

    If HechizoCasteado Then

        With UserList(Userindex)
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            ' Bonificaciones para druida
            If .Clase = eClass.Druid Then

                ' Solo con flauta magica
                If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                    If Hechizos(SpellIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        
                    ElseIf SpellIndex <> APOCALIPSIS_SPELL_INDEX Then
                        ' 10% menos de mana para todo menos apoca y descarga
                        ManaRequerida = ManaRequerida * 0.9

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)
            Call WriteUpdateUserStats(.flags.TargetUser)
            .flags.TargetUser = 0

        End With

    End If

End Sub

Sub HandleHechizoNPC(ByVal Userindex As Integer, ByVal HechizoIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
    '17/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
    '12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
    '***************************************************
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Long
    
    With UserList(Userindex)
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
            
        End If
        
        Select Case Hechizos(HechizoIndex).Tipo

            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, Userindex)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, Userindex, HechizoCasteado)

        End Select
        
        If HechizoCasteado Then
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
            
            ' Bonificacion para druidas.
            If .Clase = eClass.Druid Then
                ' Se mostro como usuario, puede ser atacado por npcs
                .flags.Ignorado = False
                
                ' Solo con flauta equipada
                If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                    If Hechizos(HechizoIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        ' Sera ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                        .flags.Ignorado = True
                    Else

                        ' 10% menos de mana para hechizos
                        If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                            ManaRequerida = ManaRequerida * 0.9

                        End If

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)
            .flags.TargetNPC = 0

        End If

    End With

End Sub

Sub LanzarHechizo(ByVal SpellIndex As Integer, ByVal Userindex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 15/03/2020
    '24/01/2007 ZaMa - Optimizacion de codigo.
    '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
    '15/03/2020: WyroX - Remuevo los chequeos de distancia, porque ya se comprueba si lanzo a un tile que ve
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex)
    
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(Userindex, "No puedes lanzar hechizos si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If PuedeLanzar(Userindex, SpellIndex) Then

            Select Case Hechizos(SpellIndex).Target

                Case TargetType.uUsuarios
                
                    #If ProteccionGM = 1 Then
                        ' WyroX: A pedido de la gente, desactivo que los GMs puedan lanzar hechizos a users y npcs
                        If (.flags.Privilegios And PlayerType.User) = 0 Then
                            Call WriteConsoleMsg(Userindex, "Los GMs no pueden lanzar hechizos a usuarios o NPCs.", FONTTYPE_SERVER)
                            Exit Sub
                        End If
                    #End If

                    If .flags.TargetUser > 0 Then
                        Call HandleHechizoUsuario(Userindex, SpellIndex)

                    Else
                        Call WriteConsoleMsg(Userindex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
                Case TargetType.uNPC
                
                    #If ProteccionGM = 1 Then
                        ' WyroX: A pedido de la gente, desactivo que los GMs puedan lanzar hechizos a users y npcs
                        If (.flags.Privilegios And PlayerType.User) = 0 Then
                            Call WriteConsoleMsg(Userindex, "Los GMs no pueden lanzar hechizos a usuarios o NPCs.", FONTTYPE_SERVER)
                            Exit Sub
                        End If
                    #End If

                    If .flags.TargetNPC > 0 Then
                        Call HandleHechizoNPC(Userindex, SpellIndex)

                    Else
                        Call WriteConsoleMsg(Userindex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
                Case TargetType.uUsuariosYnpc
                
                    #If ProteccionGM = 1 Then
                        ' WyroX: A pedido de la gente, desactivo que los GMs puedan lanzar hechizos a users y npcs
                        If (.flags.Privilegios And PlayerType.User) = 0 Then
                            Call WriteConsoleMsg(Userindex, "Los GMs no pueden lanzar hechizos a usuarios o NPCs.", FONTTYPE_SERVER)
                            Exit Sub
                        End If
                    #End If

                    If .flags.TargetUser > 0 Then
                        Call HandleHechizoUsuario(Userindex, SpellIndex)

                    ElseIf .flags.TargetNPC > 0 Then
                        Call HandleHechizoNPC(Userindex, SpellIndex)

                    Else
                        Call WriteConsoleMsg(Userindex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
                Case TargetType.uTerreno
                    Call HandleHechizoTerreno(Userindex, SpellIndex)

            End Select
        
        End If
    
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

    Exit Sub

errHandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.description & " Hechizo: " & SpellIndex & "(" & SpellIndex & "). Casteado por: " & UserList(Userindex).Name & "(" & Userindex & ").")
    
End Sub

Sub HechizoEstadoUsuario(ByVal Userindex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 03/02/2020
    'Handles the Spells that afect the Stats of an User
    '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
    '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
    '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
    '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
    '06/28/2008 NicoNZ - Agregue que se le de valor al flag Inmovilizado.
    '17/11/2008: NicoNZ - Agregado para quitar la penalizacion de vida en el ring y cambio de ecuacion.
    '13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '16/09/2010: ZaMa - Solo se hace invi para los clientes si no esta navegando.
    '03/02/2020: WyroX - Anillos anti-efectos
    '***************************************************

    Dim HechizoIndex As Integer
    Dim targetIndex  As Integer

    With UserList(Userindex)
        HechizoIndex = .flags.Hechizo
        targetIndex = .flags.TargetUser
    
        ' <-------- Agrega Invisibilidad ---------->
        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
        
            If UserList(targetIndex).Counters.Saliendo Then
                If Userindex <> targetIndex Then
                    Call WriteConsoleMsg(Userindex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
                    HechizoCasteado = False
                    Exit Sub

                End If

            End If
        
            'No usar invi mapas InviSinEfecto
            If MapInfo(UserList(targetIndex).Pos.Map).InviSinEfecto > 0 Then
                Call WriteConsoleMsg(Userindex, "La invisibilidad no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If Not EsGm(Userindex) And EsGm(targetIndex) Then
                HechizoCasteado = False
                Exit Sub
            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(Userindex, targetIndex, True)

            If Not HechizoCasteado Then Exit Sub

            UserList(targetIndex).flags.invisible = 1
        
            ' Solo se hace invi para los clientes si no esta navegando
            If UserList(targetIndex).flags.Navegando = 0 Then
                Call SetInvisible(targetIndex, UserList(targetIndex).Char.CharIndex, True)

            End If
        
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Mimetismo ---------->
        If Hechizos(HechizoIndex).Mimetiza = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
                Exit Sub

            End If
        
            If UserList(targetIndex).flags.Navegando = 1 Then
                Exit Sub

            End If

            If .flags.Navegando = 1 Then
                Exit Sub

            End If
        
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes mimetizar a un Game Master.", FontTypeNames.FONTTYPE_FIGHT)
                HechizoCasteado = False
                Exit Sub
            End If
        
            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(Userindex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If .flags.AdminInvisible = 1 Then Exit Sub
        
            'copio el char original al mimetizado
        
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
            .flags.Mimetizado = 1
        
            'ahora pongo local el del enemigo
            .Char.body = UserList(targetIndex).Char.body
            .Char.Head = UserList(targetIndex).Char.Head
            .Char.CascoAnim = UserList(targetIndex).Char.CascoAnim
            .Char.ShieldAnim = UserList(targetIndex).Char.ShieldAnim
            .Char.WeaponAnim = UserList(targetIndex).Char.WeaponAnim
        
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
       
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Envenenamiento ---------->
        If Hechizos(HechizoIndex).Envenena = 1 Then
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Sub
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If

            UserList(targetIndex).flags.Envenenado = 1
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Cura Envenenamiento ---------->
        If Hechizos(HechizoIndex).CuraVeneno = 1 Then
    
            'Verificamos que el usuario no este muerto
            If UserList(targetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(Userindex, targetIndex)

            If Not HechizoCasteado Then Exit Sub
            
            UserList(targetIndex).flags.Envenenado = 0
            
            Call InfoHechizo(Userindex)
            
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Maldicion ---------->
        If Hechizos(HechizoIndex).Maldicion = 1 Then
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Sub
            
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If

            UserList(targetIndex).flags.Maldicion = 1
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Remueve Maldicion ---------->
        If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
            UserList(targetIndex).flags.Maldicion = 0
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Bendicion ---------->
        If Hechizos(HechizoIndex).Bendicion = 1 Then
            UserList(targetIndex).flags.Bendicion = 1
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If UserList(targetIndex).flags.Paralizado = 0 Then
                If Not PuedeAtacar(Userindex, targetIndex) Then Exit Sub
            
                If Userindex <> targetIndex Then
                    Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

                End If
            
                Call InfoHechizo(Userindex)
                HechizoCasteado = True

                If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideParalizar Then
                        Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteConsoleMsg(Userindex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
            
                If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(targetIndex).flags.Inmovilizado = 1
                UserList(targetIndex).flags.Paralizado = 1
                UserList(targetIndex).Counters.Paralisis = IntervaloParalizado
            
                UserList(targetIndex).flags.ParalizedByIndex = Userindex
                UserList(targetIndex).flags.ParalizedBy = UserList(Userindex).Name
            
                Call WriteParalizeOK(targetIndex)

            End If

        End If
    
        ' <-------- Remueve Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
            ' Remueve si esta en ese estado
            If UserList(targetIndex).flags.Paralizado = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(Userindex, targetIndex, True)

                If Not HechizoCasteado Then Exit Sub
            
                Call RemoveParalisis(targetIndex)
                Call InfoHechizo(Userindex)
        
            End If

        End If
    
        ' <-------- Remueve Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
            ' Remueve si esta en ese estado
            If UserList(targetIndex).flags.Estupidez = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(Userindex, targetIndex)

                If Not HechizoCasteado Then Exit Sub
        
                UserList(targetIndex).flags.Estupidez = 0
            
                'no need to crypt this
                Call WriteDumbNoMore(targetIndex)
                Call InfoHechizo(Userindex)
        
            End If

        End If
    
        ' <-------- Revive ---------->
        If Hechizos(HechizoIndex).Revivir = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
            
                'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
                If UserList(targetIndex).flags.SeguroResu Then
                    Call WriteConsoleMsg(Userindex, "El espiritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
        
                'No usar resu en mapas con ResuSinEfecto
                If MapInfo(UserList(targetIndex).Pos.Map).ResuSinEfecto > 0 Then
                    Call WriteConsoleMsg(Userindex, "Revivir no esta permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
            
                'No podemos resucitar si nuestra barra de energia no esta llena. (GD: 29/04/07)
                If .Stats.MaxSta <> .Stats.MinSta Then
                    Call WriteConsoleMsg(Userindex, "No puedes resucitar si no tienes tu barra de energia llena.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
            
                'revisamos si necesita vara
                If .clase = eClass.Mage Then
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                            Call WriteConsoleMsg(Userindex, "Necesitas un baculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub

                        End If

                    End If

                ElseIf .Clase = eClass.Bard Then

                    If .Invent.AnilloEqpObjIndex <> LAUDELFICO And .Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                        Call WriteConsoleMsg(Userindex, "Necesitas un instrumento magico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub

                    End If

                ElseIf .Clase = eClass.Druid Then

                    If .Invent.AnilloEqpObjIndex <> FLAUTAELFICA And .Invent.AnilloEqpObjIndex <> FLAUTAMAGICA Then
                        Call WriteConsoleMsg(Userindex, "Necesitas un instrumento magico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
            
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(Userindex, targetIndex, True)

                If Not HechizoCasteado Then Exit Sub
    
                Dim EraCriminal As Boolean

                EraCriminal = criminal(Userindex)
            
                If Not criminal(targetIndex) Then
                    If targetIndex <> Userindex Then
                        .Reputacion.NobleRep = .Reputacion.NobleRep + 500

                        If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                        Call WriteConsoleMsg(Userindex, "Los Dioses te sonrien, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
                If EraCriminal And Not criminal(Userindex) Then
                    Call RefreshCharStatus(Userindex)

                End If
            
                With UserList(targetIndex)
                    'Pablo Toxic Waste (GD: 29/04/07)
                    .Stats.MinAGU = 0
                    .flags.Sed = 1
                    .Stats.MinHam = 0
                    .flags.Hambre = 1
                    Call WriteUpdateHungerAndThirst(targetIndex)
                    Call InfoHechizo(Userindex)
                    .Stats.MinMAN = 0
                    .Stats.MinSta = 0

                End With
            
                'Agregado para quitar la penalizacion de vida en el ring y cambio de ecuacion. (NicoNZ)
                If (TriggerZonaPelea(Userindex, targetIndex) <> TRIGGER6_PERMITE) Then

                    'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                    If .flags.Privilegios And PlayerType.User Then
                        .Stats.MinHp = .Stats.MinHp * (1 - UserList(targetIndex).Stats.ELV * 0.015)

                    End If

                End If
            
                If (.Stats.MinHp <= 0) Then
                    Call UserDie(Userindex)
                    Call WriteConsoleMsg(Userindex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                Else
                    Call WriteConsoleMsg(Userindex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = True

                End If
            
                If UserList(targetIndex).flags.Traveling = 1 Then
                    UserList(targetIndex).Counters.goHome = 0
                    UserList(targetIndex).flags.Traveling = 0
                    'Call WriteConsoleMsg(TargetIndex, "Tu viaje ha sido cancelado.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteMultiMessage(targetIndex, eMessages.CancelHome)

                End If
            
                Call RevivirUsuario(targetIndex)
            Else
                HechizoCasteado = False

            End If
    
        End If
    
        ' <-------- Agrega Ceguera ---------->
        If Hechizos(HechizoIndex).Ceguera = 1 Then
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Sub

            If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideCegar Then
                    Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(Userindex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If

            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If

            UserList(targetIndex).flags.Ceguera = 1
            UserList(targetIndex).Counters.Ceguera = IntervaloParalizado / 3
    
            Call WriteBlind(targetIndex)
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).Estupidez = 1 Then
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(Userindex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Sub
                                                                                                                                
            If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideAturdir Then
                    Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(Userindex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If

            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If

            If UserList(targetIndex).flags.Estupidez = 0 Then
                UserList(targetIndex).flags.Estupidez = 1
                UserList(targetIndex).Counters.Ceguera = IntervaloParalizado

            End If

            Call WriteDumb(targetIndex)
    
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End If

    End With

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, _
                     ByVal SpellIndex As Integer, _
                     ByRef HechizoCasteado As Boolean, _
                     ByVal Userindex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 07/07/2008
    'Handles the Spells that afect the Stats of an NPC
    '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
    'removidos por users de su misma faccion.
    '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
    '***************************************************

    With Npclist(NpcIndex)

        If Hechizos(SpellIndex).Invisibilidad = 1 Then
            Call InfoHechizo(Userindex)
            .flags.invisible = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Envenena = 1 Then
            If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, Userindex)
            Call InfoHechizo(Userindex)
            .flags.Envenenado = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).CuraVeneno = 1 Then
            Call InfoHechizo(Userindex)
            .flags.Envenenado = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Maldicion = 1 Then
            If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, Userindex)
            Call InfoHechizo(Userindex)
            .flags.Maldicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
            Call InfoHechizo(Userindex)
            .flags.Maldicion = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Bendicion = 1 Then
            Call InfoHechizo(Userindex)
            .flags.Bendicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Paraliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).TileExit.Map > 0 Then
                    If Not EsGm(Userindex) Then
                        Call WriteConsoleMsg(Userindex, "No puedes paralizar criaturas en esa posicion.", FontTypeNames.FONTTYPE_INFOBOLD)   '"El NPC es inmune al hechizo."
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
                                                                                                                      
                If Not PuedeAtacarNPC(Userindex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub

                End If

                Call NPCAtacado(NpcIndex, Userindex)
                Call InfoHechizo(Userindex)
                .flags.Paralizado = 1
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                HechizoCasteado = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(Userindex, eMessages.NpcInmune)
                HechizoCasteado = False
                Exit Sub

            End If

        End If
    
        If Hechizos(SpellIndex).RemoverParalisis = 1 Then
            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                If .MaestroUser = Userindex Then
                    Call InfoHechizo(Userindex)
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                    HechizoCasteado = True
                Else

                    If .NPCtype = eNPCType.GuardiaReal Then
                        If esArmada(Userindex) Then
                            Call InfoHechizo(Userindex)
                            .flags.Paralizado = 0
                            .Contadores.Paralisis = 0
                            HechizoCasteado = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(Userindex, "Solo puedes remover la paralisis de los Guardias si perteneces a su faccion.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub

                        End If
                    
                        Call WriteConsoleMsg(Userindex, "Solo puedes remover la paralisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    Else

                        If .NPCtype = eNPCType.Guardiascaos Then
                            If esCaos(Userindex) Then
                                Call InfoHechizo(Userindex)
                                .flags.Paralizado = 0
                                .Contadores.Paralisis = 0
                                HechizoCasteado = True
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(Userindex, "Solo puedes remover la paralisis de los Guardias si perteneces a su faccion.", FontTypeNames.FONTTYPE_INFO)
                                HechizoCasteado = False
                                Exit Sub

                            End If

                        End If

                    End If

                End If

            Else
                Call WriteConsoleMsg(Userindex, "Este NPC no esta paralizado", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If

        End If
     
        If Hechizos(SpellIndex).Inmoviliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNPC(Userindex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub

                End If

                With UserList(Userindex)
                '<<<< Equitando >>>
                    If .flags.Equitando = 1 Then
                        Call UnmountMontura(Userindex)
                        Call WriteEquitandoToggle(Userindex)
                        
                    End If
                End With

                If MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).TileExit.Map > 0 Then
                    If Not EsGm(Userindex) Then
                        Call WriteConsoleMsg(Userindex, "No puedes paralizar criaturas en esa posicion.", FontTypeNames.FONTTYPE_INFOBOLD)   '"El NPC es inmune al hechizo."
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
                                                                                                                                            
                Call NPCAtacado(NpcIndex, Userindex)
                .flags.Inmovilizado = 1
                .flags.Paralizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                Call InfoHechizo(Userindex)
                HechizoCasteado = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(Userindex, eMessages.NpcInmune)

            End If

        End If

    End With

    If Hechizos(SpellIndex).Mimetiza = 1 Then

        With UserList(Userindex)

            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(Userindex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If .flags.AdminInvisible = 1 Then Exit Sub
            
            If .Clase = eClass.Druid Then
                'copio el char original al mimetizado
                If .Invent.ArmourEqpObjIndex <> 0 Then
                    .CharMimetizado.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    .CharMimetizado.body = DarCuerpoDesnudo(Userindex, True) '.Char.body
                End If
                
                If .flags.Navegando <> 0 Then
                    .CharMimetizado.Head = .OrigChar.Head
                Else
                    .CharMimetizado.Head = .Char.Head
                    .CharMimetizado.body = .Char.body
                End If
                .CharMimetizado.CascoAnim = .Char.CascoAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
                .flags.Mimetizado = 1
            
                'ahora pongo lo del NPC.
                .Char.body = Npclist(NpcIndex).Char.body
                .Char.Head = Npclist(NpcIndex).Char.Head
                .Char.CascoAnim = NingunCasco
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
        
                Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            
            Else
                Call WriteConsoleMsg(Userindex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
            Call InfoHechizo(Userindex)
            HechizoCasteado = True

        End With

    End If

End Sub

Sub HechizoPropNPC(ByVal SpellIndex As Integer, _
                   ByVal NpcIndex As Integer, _
                   ByVal Userindex As Integer, _
                   ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    'Handles the Spells that afect the Life NPC
    '14/08/2007 Pablo (ToxicWaste) - Orden general.
    '18/09/2010: ZaMa - Ahora valida si podes ayudar a un npc.
    '06/04/2020: FrankoH298 - Si le lanza un hechizo al npc lo desmonta.
    '***************************************************

    Dim dano As Long

    With Npclist(NpcIndex)
    
        Dim tempX, tempY As Integer

        tempX = .Pos.X
        tempY = .Pos.Y
        'Salud
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            HechizoCasteado = CanSupportNpc(Userindex, NpcIndex)
        
            If HechizoCasteado Then
                dano = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                dano = dano + Porcentaje(dano, 3 * UserList(Userindex).Stats.ELV)
            
                Call InfoHechizo(Userindex)
                .Stats.MinHp = .Stats.MinHp + dano

                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(Userindex, "Has curado " & dano & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                
            End If
        
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then

            If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            With UserList(Userindex)
                '<<<< Equitando >>>
                If .flags.Equitando = 1 Then
                    Call UnmountMontura(Userindex)
                    Call WriteEquitandoToggle(Userindex)
                    
                End If
            End With

            Call NPCAtacado(NpcIndex, Userindex)
            dano = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            dano = dano + Porcentaje(dano, 3 * UserList(Userindex).Stats.ELV)
    
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(Userindex).Clase = eClass.Mage Then
                    If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                        dano = (dano * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                        'Aumenta dano segun el staff-
                        'Dano = (Dano* (70 + BonifBaculo)) / 100
                    Else
                        dano = dano * 0.7 'Baja dano a 70% del original

                    End If

                End If

            End If

            If UserList(Userindex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(Userindex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                dano = dano * 1.04  'laud magico de los bardos

            End If
    
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
        
            If .flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.X, .Pos.Y))

            End If
        
            'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
            dano = dano - .Stats.defM

            If dano < 0 Then dano = 0
        
            .Stats.MinHp = .Stats.MinHp - dano
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
            'Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de vida a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteMultiMessage(Userindex, eMessages.UserHitNPC, dano)
            Call CalcularDarExp(Userindex, NpcIndex, dano)
    
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                Call MuereNpc(NpcIndex, Userindex)
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFXtoMap(Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops, tempX, tempY))

            End If

        End If

    End With

End Sub

Sub InfoHechizo(ByVal Userindex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2009
    '25/07/2009: ZaMa - Code improvements.
    '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
    '***************************************************
    Dim SpellIndex As Integer
    Dim tUser      As Integer
    Dim tNPC       As Integer
    Dim tempData   As String
    
    With UserList(Userindex)
        SpellIndex = .flags.Hechizo
        tUser = .flags.TargetUser
        tNPC = .flags.TargetNPC
     
        Call DecirPalabrasMagicas(SpellIndex, Userindex)
     
        If tUser > 0 Then

            ' Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible = 1 And Userindex = tUser Then
                
                tempData = PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops)
                Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(tempData)
                
                tempData = PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)
                Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(tempData)

            Else
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If

        ElseIf tNPC > 0 Then
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessageCreateFX(Npclist(tNPC).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, Npclist(tNPC).Pos.X, Npclist(tNPC).Pos.Y))

        End If
     
        If tUser > 0 Then
            If Userindex <> tUser Then
                If .showName Then
                    Call WriteMultiMessage(Userindex, eMessages.Hechizo_HechiceroMSG_NOMBRE, SpellIndex, , , UserList(tUser).Name)
                Else
                    Call WriteMultiMessage(Userindex, eMessages.Hechizo_HechiceroMSG_ALGUIEN, SpellIndex)

                End If

                Call WriteMultiMessage(tUser, eMessages.Hechizo_TargetMSG, SpellIndex, , , .Name)
            Else
                Call WriteMultiMessage(Userindex, eMessages.Hechizo_PropioMSG, SpellIndex)

            End If

        ElseIf tNPC > 0 Then
            Call WriteMultiMessage(Userindex, eMessages.Hechizo_HechiceroMSG_CRIATURA, SpellIndex)

        End If

    End With
 
End Sub

Public Function HechizoPropUsuario(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '06/04/2020: FrankoH298 - Si le lanza un hechizo a un usuario lo desmonta.
    '***************************************************

    Dim SpellIndex  As Integer

    Dim dano        As Long

    Dim targetIndex As Integer

    SpellIndex = UserList(Userindex).flags.Hechizo
    targetIndex = UserList(Userindex).flags.TargetUser
      
    With UserList(targetIndex)

        If .flags.Muerto Then
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            Exit Function

        End If
        
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(targetIndex)
            Call WriteEquitandoToggle(targetIndex)
            
        End If

        ' <-------- Aumenta Hambre ---------->
        If Hechizos(SpellIndex).SubeHam = 1 Then
        
            Call InfoHechizo(Userindex)
        
            dano = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam + dano

            If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has restaurado " & dano & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha restaurado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has restaurado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)
    
            ' <-------- Quita Hambre ---------->
        ElseIf Hechizos(SpellIndex).SubeHam = 2 Then

            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)
            Else
                Exit Function

            End If
        
            Call InfoHechizo(Userindex)
        
            dano = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam - dano
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has quitado " & dano & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha quitado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has quitado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            If .Stats.MinHam < 1 Then
                .Stats.MinHam = 0
                .flags.Hambre = 1

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)

        End If
    
        ' <-------- Aumenta Sed ---------->
        If Hechizos(SpellIndex).SubeSed = 1 Then
        
            Call InfoHechizo(Userindex)
        
            dano = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU + dano

            If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
        
            Call WriteUpdateHungerAndThirst(targetIndex)
             
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has restaurado " & dano & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha restaurado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has restaurado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
            ' <-------- Quita Sed ---------->
        ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
        
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            dano = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU - dano
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has quitado " & dano & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha quitado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has quitado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            If .Stats.MinAGU < 1 Then
                .Stats.MinAGU = 0
                .flags.Sed = 1

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)
        
        End If
    
        ' <-------- Aumenta Agilidad ---------->
        If Hechizos(SpellIndex).SubeAgilidad = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(Userindex, targetIndex) Then Exit Function
        
            Call InfoHechizo(Userindex)
            dano = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
        
            .flags.DuracionEfecto = 1200
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + dano

            If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateDexterity(targetIndex)
    
            ' <-------- Quita Agilidad ---------->
        ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
        
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            .flags.TomoPocion = True
            dano = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - dano

            If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        
            Call WriteUpdateDexterity(targetIndex)

        End If
    
        ' <-------- Aumenta Fuerza ---------->
        If Hechizos(SpellIndex).SubeFuerza = 1 Then
    
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(Userindex, targetIndex) Then Exit Function
        
            Call InfoHechizo(Userindex)
            dano = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
        
            .flags.DuracionEfecto = 1200
    
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + dano

            If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateStrenght(targetIndex)
    
            ' <-------- Quita Fuerza ---------->
        ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
    
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            .flags.TomoPocion = True
        
            dano = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - dano

            If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
            Call WriteUpdateStrenght(targetIndex)

        End If
    
        ' <-------- Cura salud ---------->
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            'Verifica que el usuario no este muerto
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Function

            End If
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(Userindex, targetIndex) Then Exit Function
           
            dano = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            dano = dano + Porcentaje(dano, 3 * UserList(Userindex).Stats.ELV)
        
            Call InfoHechizo(Userindex)
    
            .Stats.MinHp = .Stats.MinHp + dano

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        
            Call WriteUpdateHP(targetIndex)
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has restaurado " & dano & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha restaurado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                Call SendData(SendTarget.ToPCArea, targetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                
            Else
                Call WriteConsoleMsg(Userindex, "Te has restaurado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))

            End If
        
            ' <-------- Quita salud (Dana) ---------->
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        
            If Userindex = targetIndex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function

            End If
        
            dano = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
            dano = dano + Porcentaje(dano, 3 * UserList(Userindex).Stats.ELV)
        
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(Userindex).Clase = eClass.Mage Then
                    If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                        dano = (dano * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                    Else
                        dano = dano * 0.7 'Baja dano a 70% del original

                    End If

                End If

            End If
        
            If UserList(Userindex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(Userindex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                dano = dano * 1.04  'laud magico de los bardos

            End If
        
            'cascos antimagia
            If (.Invent.CascoEqpObjIndex > 0) Then
                dano = dano - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
        
            'anillos
            If (.Invent.AnilloEqpObjIndex > 0) Then
                dano = dano - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

            End If
        
            If dano < 0 Then dano = 0
        
            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            .Stats.MinHp = .Stats.MinHp - dano
            
            'Renderizo el dano en render
            Call SendData(SendTarget.ToPCArea, targetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
            
            Call WriteUpdateHP(targetIndex)
        
            Call WriteConsoleMsg(Userindex, "Le has quitado " & dano & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        
            'Muere
            If .Stats.MinHp < 1 Then
        
                If .flags.AtacablePor <> Userindex Then
                    'Store it!
                    Call Statistics.StoreFrag(Userindex, targetIndex)
                    Call ContarMuerte(targetIndex, Userindex)

                End If
            
                .Stats.MinHp = 0
                Call ActStats(targetIndex, Userindex)
                Call UserDie(targetIndex, Userindex)

            End If
        
        End If
    
        ' <-------- Aumenta Mana ---------->
        If Hechizos(SpellIndex).SubeMana = 1 Then
        
            Call InfoHechizo(Userindex)
            .Stats.MinMAN = .Stats.MinMAN + dano

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
        
            Call WriteUpdateMana(targetIndex)
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has restaurado " & dano & " puntos de mana a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha restaurado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has restaurado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
            ' <-------- Quita Mana ---------->
        ElseIf Hechizos(SpellIndex).SubeMana = 2 Then

            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has quitado " & dano & " puntos de mana a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha quitado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has quitado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            .Stats.MinMAN = .Stats.MinMAN - dano

            If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
        
            Call WriteUpdateMana(targetIndex)
        
        End If
    
        ' <-------- Aumenta Stamina ---------->
        If Hechizos(SpellIndex).SubeSta = 1 Then
            Call InfoHechizo(Userindex)
            .Stats.MinSta = .Stats.MinSta + dano

            If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
        
            Call WriteUpdateSta(targetIndex)
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has restaurado " & dano & " puntos de energia a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha restaurado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has restaurado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            ' <-------- Quita Stamina ---------->
        ElseIf Hechizos(SpellIndex).SubeSta = 2 Then

            If Not PuedeAtacar(Userindex, targetIndex) Then Exit Function
        
            If Userindex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, targetIndex)

            End If
        
            Call InfoHechizo(Userindex)
        
            If Userindex <> targetIndex Then
                Call WriteConsoleMsg(Userindex, "Le has quitado " & dano & " puntos de energia a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(Userindex).Name & " te ha quitado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, "Te has quitado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            .Stats.MinSta = .Stats.MinSta - dano
        
            If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
            Call WriteUpdateSta(targetIndex)
        
        End If

    End With

    HechizoPropUsuario = True

End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, _
                               ByVal targetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/04/2010
    'Checks if caster can cast support magic on target user.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = targetIndex Then
            CanSupportUser = True
            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, targetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True
            Exit Function

        End If
     
        ' Victima criminal?
        If criminal(targetIndex) Then
        
            ' Casteador Ciuda?
            If Not criminal(CasterIndex) Then
            
                ' Armadas no pueden ayudar
                If esArmada(CasterIndex) Then
                    Call WriteConsoleMsg(CasterIndex, "Los miembros del ejercito real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If
                
                ' Si el ciuda tiene el seguro puesto no puede ayudar
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volveras criminal como ellos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else

                    ' Penalizacion
                    If DoCriminal Then
                        Call VolverCriminal(CasterIndex)
                    Else
                        Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                    End If

                End If

            End If
            
            ' Victima ciuda o army
        Else

            ' Casteador es caos? => No Pueden ayudar ciudas
            If esCaos(CasterIndex) Then
                Call WriteConsoleMsg(CasterIndex, "Los miembros de la legion oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
                
                ' Casteador ciuda/army?
            ElseIf Not criminal(CasterIndex) Then
                
                ' Esta en estado atacable?
                If UserList(targetIndex).flags.AtacablePor > 0 Then
                    
                    ' No esta atacable por el casteador?
                    If UserList(targetIndex).flags.AtacablePor <> CasterIndex Then
                    
                        ' Si es armada no puede ayudar
                        If esArmada(CasterIndex) Then
                            Call WriteConsoleMsg(CasterIndex, "Los miembros del ejercito real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function

                        End If
    
                        ' Seguro puesto?
                        If .flags.Seguro Then
                            Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function
                        Else
                            Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                        End If

                    End If

                End If
    
            End If

        End If

    End With
    
    CanSupportUser = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", TargetIndex: " & targetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, _
                       ByVal Userindex As Integer, _
                       ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Byte

    With UserList(Userindex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .Stats.UserHechizos(Slot) > 0 Then
                Call ChangeUserHechizo(Userindex, Slot, .Stats.UserHechizos(Slot))
            Else
                Call ChangeUserHechizo(Userindex, Slot, 0)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
                If .Stats.UserHechizos(LoopC) > 0 Then
                    Call ChangeUserHechizo(Userindex, LoopC, .Stats.UserHechizos(LoopC))
                Else
                    Call ChangeUserHechizo(Userindex, LoopC, 0)

                End If

            Next LoopC

        End If

    End With

End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, _
                              ByVal targetIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'Checks if caster can cast support magic on target Npc.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    Dim OwnerIndex As Integer
 
    With UserList(CasterIndex)
        
        OwnerIndex = Npclist(targetIndex).Owner
        
        ' Si no tiene dueno puede
        If OwnerIndex = 0 Then
            CanSupportNpc = True
            Exit Function

        End If
        
        ' Puede hacerlo si es su propio npc
        If CasterIndex = OwnerIndex Then
            CanSupportNpc = True
            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
            CanSupportNpc = True
            Exit Function

        End If
     
        ' Victima criminal?
        If criminal(OwnerIndex) Then

            ' Victima caos?
            If esCaos(OwnerIndex) Then

                ' Atacante caos?
                If esCaos(CasterIndex) Then
                    ' No podes ayudar a un npc de un caos si sos caos
                    Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que estan luchando contra un miembro de tu faccion.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If
        
            ' Uno es caos y el otro no, o la victima es pk, entonces puede ayudar al npc
            CanSupportNpc = True
            Exit Function
                
            ' Victima ciuda
        Else

            ' Atacante ciuda?
            If Not criminal(CasterIndex) Then

                ' Atacante armada?
                If esArmada(CasterIndex) Then

                    ' Victima armada?
                    If esArmada(OwnerIndex) Then
                        ' No podes ayudar a un npc de un armada si sos armada
                        Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que estan luchando contra un miembro de tu faccion.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                End If
                
                ' Uno es armada y el otro ciuda, o los dos ciudas, puede atacar si no tiene seguro
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar a criaturas que luchan contra ciudadanos debes sacarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                    
                    ' ayudo al npc sin seguro, se convierte en atacable
                Else
                    Call ToogleToAtackable(CasterIndex, OwnerIndex, True)
                    CanSupportNpc = True
                    Exit Function

                End If
                
            End If
            
            ' Atacante criminal y victima ciuda, entonces puede ayudar al npc
            CanSupportNpc = True
            Exit Function
            
        End If
    
    End With
    
    CanSupportNpc = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.Number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Sub ChangeUserHechizo(ByVal Userindex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Hechizo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    UserList(Userindex).Stats.UserHechizos(Slot) = Hechizo
    
    Call WriteChangeSpellSlot(Userindex, Slot)

End Sub

Public Sub DesplazarHechizo(ByVal Userindex As Integer, _
                            ByVal Dire As Integer, _
                            ByVal HechizoDesplazado As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (Dire <> 1 And Dire <> -1) Then Exit Sub
    If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    With UserList(Userindex)

        If Dire = 1 Then 'Mover arriba
            If HechizoDesplazado = 1 Then
                Call WriteConsoleMsg(Userindex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
                .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo

            End If

        Else 'mover abajo

            If HechizoDesplazado = MAXUSERHECHIZOS Then
                Call WriteConsoleMsg(Userindex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
                .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo

            End If

        End If

    End With

End Sub

Public Sub DisNobAuBan(ByVal Userindex As Integer, NoblePts As Long, BandidoPts As Long)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean

    EraCriminal = criminal(Userindex)
    
    With UserList(Userindex)

        'Si estamos en la arena no hacemos nada
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            'pierdo nobleza...
            .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts

            If .Reputacion.NobleRep < 0 Then
                .Reputacion.NobleRep = 0

            End If
            
            'gano bandido...
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts

            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            Call WriteMultiMessage(Userindex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)

            If criminal(Userindex) Then If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)

        End If
        
        If Not EraCriminal And criminal(Userindex) Then
            Call RefreshCharStatus(Userindex)

        End If

    End With

End Sub
