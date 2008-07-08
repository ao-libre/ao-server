Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(UserIndex)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
        
        Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (UserIndex)
            End If
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                'Store it!
                Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                
                Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          
            If UserList(UserIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(UserIndex).flags.Paralizado = 1
          UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          
          Call WriteParalizeOK(UserIndex)

     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal UserIndex As Integer)
On Error Resume Next
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(S, UserList(UserIndex).Char.CharIndex, vbCyan))
    Exit Sub
End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "No posees un báculo lo suficientemente poderoso para que puedas lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    PuedeLanzar = True
Else
   Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
   PuedeLanzar = False
End If
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Uknown
'Last modification: 06/15/2008 (NicoNZ)
'Sale del sub si no hay una posición valida.
'***************************************************
If UserList(UserIndex).NroMascotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(UserIndex).Pos.map).Pk = False Or MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).cant
    
    If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
            
            index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(index) = ind
            UserList(UserIndex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        Else
            Exit Sub
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
If UserList(UserIndex).flags.ModoCombate = False Then
    Call WriteConsoleMsg(UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(UserIndex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
    
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    If UserList(UserIndex).clase = eClass.Druid And UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.7
    Else
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    End If

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
If UserList(UserIndex).flags.ModoCombate = False Then
    Call WriteConsoleMsg(UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select


If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If

End Sub


Sub LanzarHechizo(index As Integer, UserIndex As Integer)

Dim uh As Integer

uh = UserList(UserIndex).Stats.UserHechizos(index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        Case TargetType.uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/28/2008
'Handles the Spells that afect the Stats of an User
'24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
'26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
'26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
'02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
'06/28/2008 NicoNZ - Agregué que se le de valor al flag Inmovilizado.
'***************************************************


Dim H As Integer, tU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tU = UserList(UserIndex).flags.TargetUser


If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(tU).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If UserList(tU).Counters.Saliendo Then
        If UserIndex <> tU Then
            Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        Else
            Call WriteConsoleMsg(UserIndex, "¡No puedes ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
            b = False
            Exit Sub
        End If
    End If
    
    'No usar invi mapas InviSinEfecto
    If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
        Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar invi a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If criminal(tU) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call VolverCriminal(UserIndex)
            End If
        End If
    End If
    
    'Si sos user, no uses este hechizo con GMS.
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If Not UserList(tU).flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
    End If
   
    UserList(tU).flags.invisible = 1
    Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(tU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(UserIndex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    'Si sos user, no uses este hechizo con GMS.
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If Not UserList(tU).flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
    End If
    
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
    
    'copio el char original al mimetizado
    
    With UserList(UserIndex)
        .CharMimetizado.body = .Char.body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.body = UserList(tU).Char.body
        .Char.Head = UserList(tU).Char.Head
        .Char.CascoAnim = UserList(tU).Char.CascoAnim
        .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
    
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(UserIndex)
   b = True
End If

If Hechizos(H).Envenena = 1 Then
    If UserIndex = tU Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    UserList(tU).flags.Envenenado = 1
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).CuraVeneno = 1 Then

    'Verificamos que el usuario no este muerto
    If UserList(tU).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar curar veneno a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If criminal(tU) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
        
    'Si sos user, no uses este hechizo con GMS.
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If Not UserList(tU).flags.Privilegios And PlayerType.User Then
            Exit Sub
        End If
    End If
        
    UserList(tU).flags.Envenenado = 0
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Maldicion = 1 Then
    If UserIndex = tU Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    UserList(tU).flags.Maldicion = 1
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(tU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(tU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
    If UserIndex = tU Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
     If UserList(tU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
            If UserIndex <> tU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
            Call InfoHechizo(UserIndex)
            b = True
            If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tU)
                Exit Sub
            End If
            
            If Hechizos(H).Inmoviliza = 1 Then UserList(tU).flags.Inmovilizado = 1
            UserList(tU).flags.Paralizado = 1
            UserList(tU).Counters.Paralisis = IntervaloParalizado
            
            Call WriteParalizeOK(tU)
            Call FlushBuffer(tU)
      
    End If
End If


If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(tU).flags.Paralizado = 1 Then
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If criminal(tU) And Not criminal(UserIndex) Then
                If esArmada(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)
                End If
            End If
        End If
        
        UserList(tU).flags.Inmovilizado = 0
        UserList(tU).flags.Paralizado = 0
        'no need to crypt this
        Call WriteParalizeOK(tU)
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If UserList(tU).flags.Estupidez = 1 Then
        'Para poder tirar remo estu a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If criminal(tU) And Not criminal(UserIndex) Then
                If esArmada(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If
            End If
        End If
    
        UserList(tU).flags.Estupidez = 0
        'no need to crypt this
        Call WriteDumbNoMore(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        
        'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
        If UserList(tU).flags.SeguroResu Then
            Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
    
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29/04/07)
        If UserList(UserIndex).Stats.MaxSta <> UserList(UserIndex).Stats.MinSta Then
            Call WriteConsoleMsg(UserIndex, "No puedes resucitar si no tienes tu barra de energía llena.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'revisamos si necesita vara
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas un mejor báculo para este hechizo", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UserList(UserIndex).clase = eClass.Bard Then
            If UserList(UserIndex).Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        ElseIf UserList(UserIndex).clase = eClass.Druid Then
            If UserList(UserIndex).Invent.AnilloEqpObjIndex <> FLAUTAMAGICA Then
                Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If criminal(tU) And Not criminal(UserIndex) Then
                If esArmada(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)
                End If
            End If
        End If

        Dim EraCriminal As Boolean
        EraCriminal = criminal(UserIndex)
        If Not criminal(tU) Then
            If tU <> UserIndex Then
                UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + 500
                If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.NobleRep = MAXREP
                Call WriteConsoleMsg(UserIndex, "¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        If EraCriminal And Not criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        End If
        
        
        'Pablo Toxic Waste (GD: 29/04/07)
        UserList(tU).Stats.MinAGU = 0
        UserList(tU).flags.Sed = 1
        UserList(tU).Stats.MinHam = 0
        UserList(tU).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(tU)
        Call InfoHechizo(UserIndex)
        UserList(tU).Stats.MinMAN = 0
        UserList(tU).Stats.MinSta = 0
        Dim aux As Double
        aux = UserList(tU).Stats.ELV / 100
        aux = UserList(UserIndex).Stats.MaxHP * aux
        'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - aux
        End If
        If (UserList(UserIndex).Stats.MinHP <= 0) Then
            Call UserDie(UserIndex)
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande", FontTypeNames.FONTTYPE_INFO)
            b = False
        Else
            Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado", FontTypeNames.FONTTYPE_INFO)
            b = True
        End If
        
        Call RevivirUsuario(tU)
    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
    If UserIndex = tU Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = IntervaloParalizado / 3

        Call WriteBlind(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
    If UserIndex = tU Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Ceguera = IntervaloParalizado
        End If
        Call WriteDumb(tU)
        Call FlushBuffer(tU)

        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 07/07/2008
'Handles the Spells that afect the Stats of an NPC
'04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
'removidos por users de su misma faccion.
'07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
'***************************************************
If Hechizos(hIndex).Invisibilidad = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.invisible = 1
    b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Envenenado = 1
    b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Envenenado = 0
    b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Maldicion = 0
    b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Bendicion = 1
    b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
    If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
        If Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
        Else
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If esArmada(UserIndex) Then
                    Call InfoHechizo(UserIndex)
                    Npclist(NpcIndex).flags.Paralizado = 0
                    Npclist(NpcIndex).Contadores.Paralisis = 0
                    b = True
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                
                Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If esCaos(UserIndex) Then
                        Call InfoHechizo(UserIndex)
                        Npclist(NpcIndex).flags.Paralizado = 0
                        Npclist(NpcIndex).Contadores.Paralisis = 0
                        b = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
        End If
   Else
      Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
      b = False
      Exit Sub
   End If
End If
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

If Hechizos(hIndex).Mimetiza = 1 Then
    
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
    
        
    If UserList(UserIndex).clase = eClass.Druid Then
        'copio el char original al mimetizado
        With UserList(UserIndex)
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
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
        
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End With
    Else
        Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

   Call InfoHechizo(UserIndex)
   b = True
End If
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 14/08/2007
'Handles the Spells that afect the Life NPC
'14/08/2007 Pablo (ToxicWaste) - Orden general.
'***************************************************

Dim daño As Long

'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    b = True
    
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)

    If Hechizos(hIndex).StaffAffected Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta daño segun el staff-
                'Daño = (Daño* (70 + BonifBáculo)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        daño = daño * 1.04  'laud magico de los bardos
    End If

    Call InfoHechizo(UserIndex)
    b = True
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    End If
    
    'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
    daño = daño - Npclist(NpcIndex).Stats.defM
    If daño < 0 Then daño = 0
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    Call CalcularDarExp(UserIndex, NpcIndex, daño)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(H).WAV, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(H).WAV, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.X, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            If UserList(UserIndex).showName Then
                Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, UserList(UserIndex).name & " " & Hechizos(H).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, Hechizos(H).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call WriteConsoleMsg(UserIndex, Hechizos(H).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/01/2008
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'***************************************************

Dim H As Integer
Dim daño As Integer
Dim tempChr As Integer
    
    
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser
      
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If UserIndex <> tempChr Then
      Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
      Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
      Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    'Para poder tirar cl a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    'Para poder tirar fuerza a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    
    'Verifica que el usuario no este muerto
    If UserList(tempChr).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar curar a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
       
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    If Hechizos(H).StaffAffected Then
        If UserList(UserIndex).clase = eClass.Mage Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
    End If
    
    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
    Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        'Store it!
        Call Statistics.StoreFrag(UserIndex, tempChr)
        
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If

Call FlushBuffer(tempChr)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    
    Call WriteChangeSpellSlot(UserIndex, Slot)

Else

    Call WriteChangeSpellSlot(UserIndex, Slot)

End If


End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1
        End If
    End If
End If
End Sub


Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    'Si estamos en la arena no hacemos nada
    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
    
If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts
    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + BandidoPts
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    Call WriteNobilityLost(UserIndex)
    If criminal(UserIndex) Then If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If
    
    If Not EraCriminal And criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub
