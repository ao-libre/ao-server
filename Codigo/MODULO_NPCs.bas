Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

Dim i As Integer
UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
     UserList(UserIndex).MascotasIndex(i) = 0
     UserList(UserIndex).MascotasType(i) = 0
     Exit For
  End If
Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'********************************************************
'Author: Unknown
'Llamado cuando la vida de un NPC llega a cero.
'Last Modify Date: 24/01/2007
'22/06/06: (Nacho) Chequeamos si es pretoriano
'24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
'********************************************************
On Error GoTo Errhandler
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Dim EraCriminal As Boolean
   
    If (esPretoriano(NpcIndex) = 4) Then
        'Solo nos importa si fue matado en el mapa pretoriano.
        If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
            'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
            Dim i As Integer
            Dim j As Integer
            Dim NPCI As Integer
        
            For i = 8 To 90
                For j = 8 To 90
                
                    NPCI = MapData(Npclist(NpcIndex).Pos.map, i, j).NpcIndex
                    If NPCI > 0 Then
                        If esPretoriano(NPCI) > 0 Then
                            Npclist(NPCI).Invent.ArmourEqpSlot = IIf(Npclist(NpcIndex).Pos.X > 50, 1, 5)
                        End If
                    End If
                Next j
            Next i
            Call CrearClanPretoriano(Npclist(NpcIndex).Pos.X)
        End If
    ElseIf esPretoriano(NpcIndex) > 0 Then
        If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
            Npclist(NpcIndex).Invent.ArmourEqpSlot = 0
            pretorianosVivos = pretorianosVivos - 1
        End If
    End If
   
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
        End If
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMascotas > 0 Then
            Dim T As Integer
            For T = 1 To MAXMASCOTAS
                  If UserList(UserIndex).MascotasIndex(T) > 0 Then
                      If Npclist(UserList(UserIndex).MascotasIndex(T)).TargetNPC = NpcIndex Then
                              Call FollowAmo(UserList(UserIndex).MascotasIndex(T))
                      End If
                  End If
            Next T
        End If
        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then
            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.map, MiNPC.Pos.X, MiNPC.Pos.Y)
            Else
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                    UserList(UserIndex).Stats.Exp = MAXEXP
                Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            MiNPC.flags.ExpCount = 0
        End If
        
        '[/KEVIN]
        Call WriteConsoleMsg(UserIndex, "Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then _
            UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
        
        EraCriminal = criminal(UserIndex)
        
        If MiNPC.Stats.Alineacion = 0 Then
            If MiNPC.Numero = Guardias Then
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
            If MiNPC.MaestroUser = 0 Then
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + vlASESINO
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
        ElseIf MiNPC.Stats.Alineacion = 1 Then
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 2 Then
            UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlASESINO / 2
            If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then _
                UserList(UserIndex).Reputacion.NobleRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 4 Then
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        End If
        If criminal(UserIndex) And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
        If Not criminal(UserIndex) And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
        
        If EraCriminal And Not criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        ElseIf Not EraCriminal And criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        End If
        
        Call CheckUserLevel(UserIndex)
    End If ' Userindex > 0
   
    If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)
    End If
   
    
    
Exit Sub

Errhandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .Attacking = 0
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UseAINow = False
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
        .AIPersonalidad = e_Personalidad.ninguna
    End With
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = vbNullString
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones
    Npclist(NpcIndex).Expresiones(j) = vbNullString
Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).GiveGLD = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0
    
    If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc)
    
    Npclist(NpcIndex).MaestroUser = 0
    Npclist(NpcIndex).MaestroNpc = 0
    
    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.map = 0
    Npclist(NpcIndex).Orig.X = 0
    Npclist(NpcIndex).Orig.Y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).Pos.map = 0
    Npclist(NpcIndex).Pos.X = 0
    Npclist(NpcIndex).Pos.Y = 0
    Npclist(NpcIndex).SkillDomar = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNPC = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).desc = vbNullString
    
    
    Dim j As Integer
    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo Errhandler

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
        Call EraseNPCChar(NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

Exit Sub

Errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
    
    If LegalPos(Pos.map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 3 And _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 2 And _
        MapData(Pos.map, Pos.X, Pos.Y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Sub
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.map, OrigPos.X, OrigPos.Y) Then
        
        map = OrigPos.map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.map = mapa 'mapa
        altpos.map = mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.map, newpos.X, newpos.Y, PuedeAgua) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.map = newpos.map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    map = altpos.map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.map = map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, map, nIndex, map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.map = newpos.map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.map, nIndex, newpos.map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        map = newpos.map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
    End If
    
    'Crea el NPC
    Call MakeNPCChar(True, map, nIndex, map, X, Y)

End Sub

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(map, X, Y).NpcIndex = NpcIndex
    
    If Not toMap Then
        Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.heading, Npclist(NpcIndex).Char.CharIndex, X, Y, 0, 0, 0, 0, 0, vbNullString, 0, 0)
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(NpcIndex)
    End If
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading)
    If NpcIndex > 0 Then
        Npclist(NpcIndex).Char.body = body
        Npclist(NpcIndex).Char.Head = Head
        Npclist(NpcIndex).Char.heading = heading
        
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, Head, heading, Npclist(NpcIndex).Char.CharIndex, 0, 0, 0, 0, 0))
    End If
End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex))

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
        ' es una posicion legal
        If LegalPos(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y) Then Exit Sub
            

            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(Npclist(NpcIndex).Char.CharIndex, nPos.X, nPos.Y))
            
            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).Char.heading = nHeading
            MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        End If
Else ' No es mascota
        ' Controlamos que la posicion sea legal, los npc que
        ' no son mascotas tienen mas restricciones de movimiento.
        If LegalPosNPC(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
            
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.map, nPos.X, nPos.Y) Then Exit Sub
            
            '[Alejo-18-5]
            'server

            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(Npclist(NpcIndex).Char.CharIndex, nPos.X, nPos.Y))

            
            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).Char.heading = nHeading
            MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        Else
            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        
        End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo Errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
Errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

Dim N As Integer
N = RandomNumber(1, 100)
If N < 30 Then
    UserList(UserIndex).flags.Envenenado = 1
    Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/15/2008
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'06/15/2008 -> Optimizé el codigo. (NicoNZ)
'***************************************************
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer
'Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

PuedeAgua = Npclist(nIndex).flags.AguaValida
'PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1

'it = 0

'Do While Not PosicionValida
        
        Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
        Call ClosestLegalPos(Pos, altpos, PuedeAgua)
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida

        If newpos.X <> 0 And newpos.Y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.map = newpos.map
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            If altpos.X <> 0 And altpos.Y <> 0 Then
                Npclist(nIndex).Pos.map = altpos.map
                Npclist(nIndex).Pos.X = altpos.X
                Npclist(nIndex).Pos.Y = altpos.Y
                PosicionValida = True
            Else
                PosicionValida = False
            End If
        End If
        
        'it = it + 1
        
        'If it > MAXSPAWNATTEMPS Then
        If Not PosicionValida Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            'Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.map & " Index:" & NpcIndex)
            Exit Function
        End If
'Loop

'asignamos las nuevas coordenas
map = newpos.map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(True, map, nIndex, map, X, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
End If

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).Pos.map = map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As npc)

'SI EL NPC TIENE ORO LO TIRAMOS
'Pablo (ToxicWaste): Ahora se puede poner más de 10k de drop de oro en los NPC.
If MiNPC.GiveGLD > 0 Then
    Dim MiObj As Obj
    Dim MiAux As Double
    MiAux = MiNPC.GiveGLD
    Do While MiAux > 10000
        MiObj.amount = 10000
        MiObj.ObjIndex = iORO
        Call TirarItemAlPiso(MiNPC.Pos, MiObj)
        MiAux = MiAux - 10000
    Loop
    If MiAux > 0 Then
        MiObj.amount = MiAux
        MiObj.ObjIndex = iORO
        Call TirarItemAlPiso(MiNPC.Pos, MiObj)
    End If
    
End If

End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer los NPCS se deberá usar la
'nueva clase clsIniReader.
'
'Alejo
'
'###################################################

Dim NpcIndex As Integer
Dim Leer As clsIniReader

Set Leer = LeerNPCs

'If requested index is invalid, abort
If Not Leer.KeyExists("NPC" & NpcNumber) Then
    OpenNPC = MAXNPCS + 1
    Exit Function
End If

NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = Leer.GetValue("NPC" & NpcNumber, "Name")
Npclist(NpcIndex).desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))

'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))

Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))


Npclist(NpcIndex).Stats.MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
Next LoopC


If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
    Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
    Next LoopC
End If



Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
If LenB(aux) = 0 Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = vbNullString
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = TipoAI.SigueAmo 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNPC = 0

End Sub

