Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1
If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
    UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(UserIndex).Counters.TiempoOculto = 0
    UserList(UserIndex).flags.Oculto = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
End If



Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(UserIndex).flags.Oculto = 1
    Suerte = (-0.000001 * (100 - Skill) ^ 3)
    Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
    Suerte = Suerte + (-0.0088 * (100 - Skill))
    Suerte = Suerte + (0.9571)
    Suerte = Suerte * IntervaloOculto
    UserList(UserIndex).Counters.TiempoOculto = Suerte
  
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).clase)

If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        '(Nacho)
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            UserList(UserIndex).Char.body = iFragataReal
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            UserList(UserIndex).Char.body = iFragataCaos
        Else
            If criminal(UserIndex) Then
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaPk
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraPk
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonPk
            Else
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
            End If
        End If
    Else
        UserList(UserIndex).Char.body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
    
Else
    
    UserList(UserIndex).flags.Navegando = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteNavigateToggle(UserIndex)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call WriteConsoleMsg(UserIndex, "No tenes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).amount
    End If
Next i

If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).amount = UserList(UserIndex).Invent.Object(i).amount - cant
        If (UserList(UserIndex).Invent.Object(i).amount <= 0) Then
            cant = Abs(UserList(UserIndex).Invent.Object(i).amount)
            UserList(UserIndex).Invent.Object(i).amount = 0
            UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        Else
            cant = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        If (cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
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


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
    End If
    
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
    If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

If CarpinteroTieneMateriales(UserIndex, ItemIndex) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(UserIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
    
    Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
    Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
    End If
    
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))


    UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
    If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End If
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
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


Public Sub DoLingotes(ByVal UserIndex As Integer)
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer

    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Invent.Object(Slot).amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount - MineralesParaLingote(obji)
    If UserList(UserIndex).Invent.Object(Slot).amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteConsoleMsg(UserIndex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFO)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End Sub

Function ModNavegacion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Pirat
        ModNavegacion = 1
    Case eClass.Fisher
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Miner
        ModFundicion = 1
    Case eClass.Blacksmith
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

Select Case clase
    Case eClass.Carpenter
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single
Select Case clase
    Case eClass.Blacksmith
        ModHerreriA = 1
    Case eClass.Miner
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer
    Select Case clase
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

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 06/18/08 (NicoNZ)
'
'***************************************************

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer


If Npclist(NpcIndex).MaestroUser = UserIndex Then
    Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "La criatura ya te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    ElseIf Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(UserIndex).Stats.UserSkills(eSkill.Domar))
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
    Else
        puntosRequeridos = Npclist(NpcIndex).flags.Domable
    End If
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
        index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(index) = NpcIndex
        UserList(UserIndex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        'Entreno domar. Es un 30% más dificil si no sos druida.
        If UserList(UserIndex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
            Call SubirSkill(UserIndex, Domar)
        End If
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
        'Entreno domar aunque no logue domar.
        If UserList(UserIndex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
            Call SubirSkill(UserIndex, Domar)
        End If
    End If

Else
    Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
            UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
            UserList(UserIndex).Counters.Mimetismo = 0
            UserList(UserIndex).flags.Mimetizado = 0
        End If
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Char.body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(map, X, Y) Then Exit Sub

With posMadera
    .map = map
    .X = X
    .Y = Y
End With

If MapData(map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).ObjInfo.amount < 3 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.amount = MapData(map, X, Y).ObjInfo.amount \ 3
    
    Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).clase = eClass.Fisher Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(UserIndex).clase = eClass.Fisher Then
        MiObj.amount = RandomNumber(1, 4)
    Else
        MiObj.amount = 1
    End If
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Pesca)

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UserList(UserIndex).clase = eClass.Fisher Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.amount = RandomNumber(1, 5)
        Else
            MiObj.amount = 1
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call SubirSkill(UserIndex, Pesca)
End If

    UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
    If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        
Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call WriteConsoleMsg(LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If


Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean
'Tiene los Guantes de Hurto equipados?
GuantesHurto = True
If UserList(LadrOnIndex).Invent.AnilloEqpObjIndex = 0 Then
    GuantesHurto = False
Else
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then GuantesHurto = False
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then GuantesHurto = False
End If


If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
        
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).clase = eClass.Thief) Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                
                If UserList(LadrOnIndex).clase = eClass.Thief Then
                ' Si no tine puestos los guantes de hurto roba un 20% menos. Pablo (ToxicWaste)
                    If GuantesHurto Then
                        N = RandomNumber(100, 1000)
                    Else
                        N = RandomNumber(80, 800)
                    End If
                Else
                    N = RandomNumber(1, 100)
                End If
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + N
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(VictimaIndex)
    End If

    If Not criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).amount Then
         num = UserList(VictimaIndex).Invent.Object(i).amount
    End If
                
    MiObj.amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).amount = UserList(VictimaIndex).Invent.Object(i).amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).clase = eClass.Thief Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(UserIndex).clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Cleric, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).clase = eClass.Assasin Then
            daño = Round(daño * 1.4, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.5)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
    End If
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).clase = eClass.Lumberjack Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(UserIndex).clase = eClass.Lumberjack Then
        MiObj.amount = RandomNumber(1, 4)
    Else
        MiObj.amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Talar)

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(UserIndex).clase = eClass.Miner Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UserList(UserIndex).clase = eClass.Miner Then
        MiObj.amount = RandomNumber(1, 6) '(NicoNZ) 04/25/2008
    Else
        MiObj.amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Mineria)

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If
    
If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMeditateToggle(UserIndex)
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    
    cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, PorcentajeRecuperoMana)
    If cant <= 0 Then cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call WriteUpdateMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 28/01/2007
'Implements the pick pocket skill of the Bandit :)
'***************************************************
If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(UserIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
        Call RobarObjeto(UserIndex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
If UserList(UserIndex).clase <> eClass.Thief Then Exit Sub
    
'una vez más, la forma de reconocer los guantes es medio patética.
If UserList(UserIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then
            Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        Call FlushBuffer(VictimIndex)
    End If
End Sub

