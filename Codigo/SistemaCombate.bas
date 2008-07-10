Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18


Function ModificadorEvasion(ByVal clase As eClass) As Single

    ModificadorEvasion = ModClase(clase).Evasion

End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single

    ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
    
    ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

End Function

Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
    
    ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

End Function

Function ModicadorDañoClaseWrestling(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseWrestling = ModClase(clase).DañoWrestling

End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single

    ModEvasionDeEscudoClase = ModClase(clase).Escudo

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function


Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * _
ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
     With UserList(UserIndex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))
    End With
End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
   PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function

Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
       (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueArma(UserIndex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWrestling(UserIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(UserIndex, Proyectiles)
       Else
            Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wrestling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(UserIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Call WriteBlockedWithShieldUser(UserIndex)
                Call SubirSkill(UserIndex, Defensa)
            End If
        End If
    End If
End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long

''sacar esto si no queremos q la matadracos mate el Dragon si o si
Dim matoDragon As Boolean
matoDragon = False


If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata Dragones?
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
            Else ' Sino es Dragon daño es 1
                DañoArma = 1
                DañoMaxArma = 1
            End If
        Else ' daño comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            DañoArma = 1 ' Si usa la espada mataDragones daño es 1
            DañoMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    'Pablo (ToxicWaste)
    ModifClase = ModicadorDañoClaseWrestling(UserList(UserIndex).clase)
    DañoArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
    DañoMaxArma = 3
End If

DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)

''sacar esto si no queremos q la matadracos mate el Dragon si o si
If matoDragon Then
    CalcularDaño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDaño = ((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase
End If

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
Dim daño As Long



daño = CalcularDaño(UserIndex, NpcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then _
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

daño = daño - Npclist(NpcIndex).Stats.def

If daño < 0 Then daño = 0

'[KEVIN]
Call WriteUserHitNPC(UserIndex, daño)
Call CalcularDarExp(UserIndex, NpcIndex, daño)
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
'[/KEVIN]

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(UserIndex) Then
       Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
       Call SubirSkill(UserIndex, Apuñalar)
    End If
    'trata de dar golpe crítico
    Call DoGolpeCritico(UserIndex, NpcIndex, 0, daño)
    
End If

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        
        ' Si era un Dragon perdemos la espada mataDragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(UserIndex).name & " mató un dragón")
        End If
        
        
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(UserIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                    Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                    Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                End If
            End If
        Next j
        
        Call MuereNpc(NpcIndex, UserIndex)
End If

End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim daño As Integer, Lugar As Integer, absorbido As Integer
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData



daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = daño

If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)


Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           Dim Obj2 As ObjData
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
           Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           End If
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
End Select

Call WriteNPCHitUser(UserIndex, Lugar, daño)

If UserList(UserIndex).flags.Privilegios And PlayerType.User Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño

If UserList(UserIndex).flags.Meditando Then
    If daño > Fix(UserList(UserIndex).Stats.MinHP / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    End If
End If

'Muere el usuario
If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
    
    'Si lo mato un guardia
    If criminal(UserIndex) And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        Call RestarCriminalidad(UserIndex)
        If Not criminal(UserIndex) And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
        End If
    End If
    
    
    Call UserDie(UserIndex)

End If

End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    If UserList(UserIndex).Reputacion.BandidoRep > 0 Then
         UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO
         If UserList(UserIndex).Reputacion.BandidoRep < 0 Then UserList(UserIndex).Reputacion.BandidoRep = 0
    ElseIf UserList(UserIndex).Reputacion.LadronesRep > 0 Then
         UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - (vlCAZADOR * 10)
         If UserList(UserIndex).Reputacion.LadronesRep < 0 Then UserList(UserIndex).Reputacion.LadronesRep = 0
    End If
    
    If EraCriminal And Not criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub


Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales Or (Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
            If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
If (Not UserList(UserIndex).flags.Privilegios And PlayerType.User) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function

' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, UserIndex, False)

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex

    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
End If

If NpcImpacto(NpcIndex, UserIndex) Then
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
    If UserList(UserIndex).flags.Meditando = False Then
        If UserList(UserIndex).flags.Navegando = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))
        End If
    End If
    
    Call NpcDaño(NpcIndex, UserIndex)
    Call WriteUpdateHP(UserIndex)
    '¿Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
Else
    Call WriteNPCSwing(UserIndex)
End If



'-----Tal vez suba los skills------
Call SubirSkill(UserIndex, Tacticas)

'Controla el nivel del usuario
Call CheckUserLevel(UserIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim daño As Integer
    Dim ANpc As npc
    ANpc = Npclist(Atacante)
    
    daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
    Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño
    
    If Npclist(Victima).Stats.MinHP < 1 Then
        
        If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        If Npclist(Atacante).MaestroUser > 0 Then
            Call FollowAmo(Atacante)
        End If
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
    End If
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
If Npclist(Atacante).CanAttack = 1 Then
       Npclist(Atacante).CanAttack = 0
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
End If

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)


If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
    Exit Sub
End If

Call NPCAtacado(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    End If
    
    Call UserDañoNpc(UserIndex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call WriteUserSwing(UserIndex)
End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
'Check bow's interval
If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

'Check Spell-Magic interval
If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then
        Exit Sub
    End If
End If

'Quitamos stamina
If UserList(UserIndex).Stats.MinSta >= 10 Then
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
Else
    Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
 
'UserList(UserIndex).flags.PuedeAtacar = 0

Dim AttackPos As WorldPos
AttackPos = UserList(UserIndex).Pos
Call HeadtoPos(UserList(UserIndex).Char.heading, AttackPos)
   
'Exit if not legal
If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Exit Sub
End If
    
Dim index As Integer
index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).UserIndex
    
'Look for user
If index > 0 Then
    Call UsuarioAtacaUsuario(UserIndex, index)
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(index)
    Exit Sub
End If
    
'Look for NPC
If MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
    If Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
            
        If Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
            MapInfo(Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).Pos.map).Pk = False Then
                Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
        End If

        Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex)
            
    Else
        Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
    End If
        
    Call WriteUpdateUserStats(UserIndex)
        
    Exit Sub
End If
    
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
Call WriteUpdateUserStats(UserIndex)


If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    
If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If

'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
              
              Call WriteBlockedWithShieldOther(AtacanteIndex)
              Call WriteBlockedWithShieldUser(VictimaIndex)
              
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
Call FlushBuffer(VictimaIndex)
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).Pos.X, UserList(AtacanteIndex).Pos.Y))
    
    If UserList(VictimaIndex).flags.Navegando = 0 Then
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
    End If
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
    'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acción
    If UserList(AtacanteIndex).clase = eClass.Bandit Then Call DoHurtar(AtacanteIndex, VictimaIndex)
    'y ahora, el ladrón puede llegar a paralizar con el golpe.
    If UserList(AtacanteIndex).clase = eClass.Thief Then Call DoHandInmo(AtacanteIndex, VictimaIndex)
    
Else
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, UserList(AtacanteIndex).Pos.X, UserList(AtacanteIndex).Pos.Y))
    Call WriteUserSwing(AtacanteIndex)
    Call WriteUserAttackedSwing(VictimaIndex, AtacanteIndex)
End If

If UserList(AtacanteIndex).clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer

Dim Obj As ObjData

daño = CalcularDaño(AtacanteIndex)
antdaño = daño

Call UserEnvenena(AtacanteIndex, VictimaIndex)

If UserList(AtacanteIndex).flags.Navegando = 1 And UserList(AtacanteIndex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim Resist As Byte
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

Select Case Lugar
    Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
        Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        absorbido = absorbido + defbarco - Resist
        daño = daño - absorbido
        If daño < 0 Then daño = 1
        End If
    Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
            Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
            Dim Obj2 As ObjData
            If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
            Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
            absorbido = absorbido + defbarco - Resist
            daño = daño - absorbido
            If daño < 0 Then daño = 1
        End If
End Select

Call WriteUserHittedUser(AtacanteIndex, Lugar, UserList(VictimaIndex).Char.CharIndex, daño)
Call WriteUserHittedByUser(VictimaIndex, Lugar, UserList(AtacanteIndex).Char.CharIndex, daño)

UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño

Call SubirSkill(VictimaIndex, Tacticas)

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
        'Si usa un arma quizas suba "Combate con armas"
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil Then
                'es un Arco. Sube Armas a Distancia
                Call SubirSkill(AtacanteIndex, Proyectiles)
            Else
                'Sube combate con armas.
                Call SubirSkill(AtacanteIndex, Armas)
            End If
        Else
        'sino tal vez lucha libre
            Call SubirSkill(AtacanteIndex, Wrestling)
        End If
                
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
            Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
            Call SubirSkill(AtacanteIndex, Apuñalar)
        End If
        'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
        Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)
End If


If UserList(VictimaIndex).Stats.MinHP <= 0 Then
    'Store it!
    Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
    
    Call ContarMuerte(VictimaIndex, AtacanteIndex)
    
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then
                Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
                Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
            End If
        End If
    Next j
    
    Call ActStats(VictimaIndex, AtacanteIndex)
Else
    'Está vivo - Actualizamos el HP
    Call WriteUpdateHP(VictimaIndex)
End If

'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)

Call FlushBuffer(VictimaIndex)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************

    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If Not criminal(attackerIndex) And Not criminal(VictimIndex) Then
        Call VolverCriminal(attackerIndex)
    End If
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(VictimIndex).Char.FX = 0
        UserList(VictimIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    EraCriminal = criminal(attackerIndex)
    
    If Not criminal(VictimIndex) Then
        UserList(attackerIndex).Reputacion.BandidoRep = UserList(attackerIndex).Reputacion.BandidoRep + vlASALTO
        If UserList(attackerIndex).Reputacion.BandidoRep > MAXREP Then _
            UserList(attackerIndex).Reputacion.BandidoRep = MAXREP
        UserList(attackerIndex).Reputacion.NobleRep = UserList(attackerIndex).Reputacion.NobleRep / 2
        If UserList(attackerIndex).Reputacion.NobleRep < 0 Then _
            UserList(attackerIndex).Reputacion.NobleRep = 0
    Else
        UserList(attackerIndex).Reputacion.NobleRep = UserList(attackerIndex).Reputacion.NobleRep + vlNoble
        If UserList(attackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(attackerIndex).Reputacion.NobleRep = MAXREP
    End If
    
    If EraCriminal And Not criminal(attackerIndex) Then
        Call RefreshCharStatus(attackerIndex)
    ElseIf Not EraCriminal And criminal(attackerIndex) Then
        Call RefreshCharStatus(attackerIndex)
    End If

    If criminal(attackerIndex) Then If UserList(attackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(attackerIndex)
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 24/01/2007
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'***************************************************
Dim T As eTrigger6
Dim rank As Integer
'MUY importante el orden de estos "IF"...

'Estas muerto no podes atacar
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

'No podes atacar a alguien muerto
If UserList(VictimIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

'Estamos en una Arena? o un trigger zona segura?
T = TriggerZonaPelea(attackerIndex, VictimIndex)

If T = eTrigger6.TRIGGER6_PERMITE Then
    PuedeAtacar = True
    Exit Function
ElseIf T = eTrigger6.TRIGGER6_PROHIBE Then
    PuedeAtacar = False
    Exit Function
ElseIf T = eTrigger6.TRIGGER6_AUSENTE Then
    'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
    If Not UserList(VictimIndex).flags.Privilegios And PlayerType.User Then
        If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
End If

'Consejeros no pueden atacar
'If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
'    PuedeAtacar = False
'    Exit Sub
'End If

'Estas queriendo atacar a un GM?
rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Sos un Armada atacando un ciudadano?
If (Not criminal(VictimIndex)) And (esArmada(attackerIndex)) Then
    Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Sos un Caos atacando otro caos?
If esCaos(VictimIndex) And esCaos(attackerIndex) Then
    Call WriteConsoleMsg(attackerIndex, "Los miembros de la legión oscura tienen prohibido atacarse entre sí.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Tenes puesto el seguro?
If UserList(attackerIndex).flags.Seguro Then
    If Not criminal(VictimIndex) Then
        Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro ingresando /seg", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
End If

'Estas en un Mapa Seguro?
If MapInfo(UserList(VictimIndex).Pos.map).Pk = False Then
    If esArmada(attackerIndex) Then
        If UserList(attackerIndex).Faccion.RecompensasReal > 11 Then
            If UserList(VictimIndex).Pos.map = 58 Or UserList(VictimIndex).Pos.map = 59 Or UserList(VictimIndex).Pos.map = 60 Then
            Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
            Exit Function
            End If
        End If
    End If
    If esCaos(attackerIndex) Then
        If UserList(attackerIndex).Faccion.RecompensasCaos > 11 Then
            If UserList(VictimIndex).Pos.map = 151 Or UserList(VictimIndex).Pos.map = 156 Then
            Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
            Exit Function
            End If
        End If
    End If
    Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
If MapData(UserList(VictimIndex).Pos.map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
    MapData(UserList(attackerIndex).Pos.map, UserList(attackerIndex).Pos.X, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

PuedeAtacar = True

End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
'***************************************************

'Estas muerto?
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Sos consejero?
If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
    'No pueden atacar NPC los Consejeros.
    PuedeAtacarNPC = False
    Exit Function
End If

'Estas en modo Combate?
If Not UserList(attackerIndex).flags.ModoCombate Then
    Call WriteConsoleMsg(attackerIndex, "Debes estar en modo de combate poder atacar al NPC.", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Es una criatura atacable?
If Npclist(NpcIndex).Attackable = 0 Then
'No es una criatura atacable
    Call WriteConsoleMsg(attackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Es valida la distancia a la cual estamos atacando?
If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   PuedeAtacarNPC = False
   Exit Function
End If

'Es una criatura No-Hostil?
If Npclist(NpcIndex).Hostile = 0 Then
'Es una criatura No-Hostil.
    'Es Guardia del Caos?
    If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
        'Lo quiere atacar un caos?
        If esCaos(attackerIndex) Then
            Call WriteConsoleMsg(attackerIndex, "No puedes atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
    'Es guardia Real?
    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        'Lo quiere atacar un Armada?
        If esArmada(attackerIndex) Then
            Call WriteConsoleMsg(attackerIndex, "No puedes atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function
        End If
        'Tienes el seguro puesto?
        If UserList(attackerIndex).flags.Seguro Then
            Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para poder Atacar Guardias Reales utilizando /seg", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function
        Else
            Call WriteConsoleMsg(attackerIndex, "Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
            Call VolverCriminal(attackerIndex)
            PuedeAtacarNPC = True
            Exit Function
        End If
    End If

    'No era un Guardia, asi que es una criatura No-Hostil común.
    'Para asegurarnos que no sea una Mascota:
    If Npclist(NpcIndex).MaestroUser = 0 Then
        'Si sos ciudadano tenes que quitar el seguro para atacarla.
        If Not criminal(attackerIndex) Then
            'Sos ciudadano, tenes el seguro puesto?
            If UserList(attackerIndex).flags.Seguro Then
            'Tiene el seguro puesto. No puede atacar
                Call WriteConsoleMsg(attackerIndex, "Para atacar a este NPC debés quitar el seguro", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            Else
            'No tiene seguro puesto. Puede atacar pero es penalizado.
                Call WriteConsoleMsg(attackerIndex, "Atacaste un NPC No-Hostil. Continua haciendolo y serás Criminal.", FontTypeNames.FONTTYPE_INFO)
                Call DisNobAuBan(attackerIndex, 10000, 1000)
                PuedeAtacarNPC = True
                Exit Function
            End If
        End If
    End If
End If

'Es el NPC mascota de alguien?
If Npclist(NpcIndex).MaestroUser > 0 Then
    If Not criminal(Npclist(NpcIndex).MaestroUser) Then
    'Es mascota de un Ciudadano.
        If esArmada(attackerIndex) Then
        'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
            Call WriteConsoleMsg(attackerIndex, "Los Armadas no pueden atacar mascotas de Ciudadanos.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function
        End If
        If Not criminal(attackerIndex) Then
        'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
            If UserList(attackerIndex).flags.Seguro Then
            'El atacante tiene el seguro puesto. No puede atacar.
                Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de Ciudadanos debes quitar el seguro utilizando /seg", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            Else
            'El atacante no tiene el seguro puesto. Recibe penalización.
                Call WriteConsoleMsg(attackerIndex, "Has atacado la Mascota de un Ciudadano. Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                Call VolverCriminal(attackerIndex)
                PuedeAtacarNPC = True
                Exit Function
            End If
        End If
    Else
    'Es mascota de un Criminal.
        If esCaos(Npclist(NpcIndex).MaestroUser) Then
        'Es Caos el Dueño.
            If esCaos(attackerIndex) Then
            'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                Call WriteConsoleMsg(attackerIndex, "Los miembros de la Legión Oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        End If
    End If
End If

'Es el Rey Preatoriano?
If esPretoriano(NpcIndex) = 4 Then
    If pretorianosVivos > 0 Then
        Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If
End If


PuedeAtacarNPC = True

End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
Dim ExpaDar As Long

'[Nacho] Chekeamos que las variables sean validas para las operaciones
If ElDaño <= 0 Then ElDaño = 0
If Npclist(NpcIndex).Stats.MaxHP <= 0 Then Exit Sub
If ElDaño > Npclist(NpcIndex).Stats.MinHP Then ElDaño = Npclist(NpcIndex).Stats.MinHP

'[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHP))
If ExpaDar <= 0 Then Exit Sub

'[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
    ExpaDar = Npclist(NpcIndex).flags.ExpCount
    Npclist(NpcIndex).flags.ExpCount = 0
Else
    Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
End If

'[Nacho] Le damos la exp al user
If ExpaDar > 0 Then
    If UserList(UserIndex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(UserIndex, ExpaDar, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
    Else
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
        If UserList(UserIndex).Stats.Exp > MAXEXP Then _
            UserList(UserIndex).Stats.Exp = MAXEXP
        Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call CheckUserLevel(UserIndex)
End If

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End If
End If

Call FlushBuffer(VictimaIndex)
End Sub

