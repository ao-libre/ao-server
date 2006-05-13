Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 10     'Lo atamos con alambre.... en la 11.6 el sistema de ocultarse debería de estar bien hecho
End If

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50

'cazador con armadura de cazador oculto no se hace visible
If UCase$(UserList(UserIndex).Clase) = "CAZADOR" And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
    If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
        Exit Sub
    End If
End If


res = RandomNumber(1, Suerte)

If res > 9 Then
    UserList(UserIndex).flags.Oculto = 0
    If UserList(UserIndex).flags.Invisible = 0 Then
        'no hace falta encriptar este (se jode el gil que bypassea esto)
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
    End If
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 7
End If

If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res <= 5 Then
    UserList(UserIndex).flags.Oculto = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
    Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
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
ModNave = ModNavegacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Body = Barco.Ropaje
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal
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
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
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
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "NAVEG")

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
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
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - cant
        If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
            cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
            UserList(UserIndex).Invent.Object(i).Amount = 0
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
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes suficientes madera." & FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de hierro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de plata." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de oro." & FONTTYPE_INFO)
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
'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has construido el arma!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has construido el escudo!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has construido el casco!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has construido la armadura!." & FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

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
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has construido el objeto!" & FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 13
        Case iMinerales.PlataCruda
            MineralesParaLingote = 25
        Case iMinerales.OroCrudo
            MineralesParaLingote = 50
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
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)
    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has obtenido un lingote!!!" & FONTTYPE_INFO)
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has obtenido un lingote!" & FONTTYPE_INFO)
    


UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "MINERO"
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "HERRERO"
        ModHerreriA = 1
    Case "MINERO"
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
    Select Case UCase$(Clase)
        Case "DRUIDA"
            ModDomar = 6
        Case "CAZADOR"
            ModDomar = 6
        Case "CLERIGO"
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
    With UserList(UserIndex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) _
            * (.UserSkills(eSkill.Domar) / ModDomar(UserList(UserIndex).Clase)) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3)
    End With
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
'Call LogTarea("Sub DoDomar")

If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
        Dim index As Integer
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(index) = NpcIndex
        UserList(UserIndex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
        Call SubirSkill(UserIndex, Domar)
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
            UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
            UserList(UserIndex).Counters.Mimetismo = 0
            UserList(UserIndex).flags.Mimetizado = 0
        End If
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás demasiado lejos para prender la fogata." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes hacer fogatas estando muerto." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
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
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount \ 3
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    
    Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
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


If UCase$(UserList(UserIndex).Clase) = "PESCADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = 1
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Pesca)

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

If UCase(UserList(UserIndex).Clase) = "PESCADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Select Case iSkill
Case 0:         Suerte = 0
Case 1 To 10:   Suerte = 60
Case 11 To 20:  Suerte = 54
Case 21 To 30:  Suerte = 49
Case 31 To 40:  Suerte = 43
Case 41 To 50:  Suerte = 38
Case 51 To 60:  Suerte = 32
Case 61 To 70:  Suerte = 27
Case 71 To 80:  Suerte = 21
Case 81 To 90:  Suerte = 16
Case 91 To 100: Suerte = 11
Case Else:      Suerte = 0
End Select

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
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has pescado algunos peces!" & FONTTYPE_INFO)
        
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
    End If
    
    Call SubirSkill(UserIndex, Pesca)
End If

Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||Debes quitar el seguro para robar" & FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||No puedes robar a otros miembros de las fuerzas del caos" & FONTTYPE_FIGHT)
    Exit Sub
End If

If UserList(VictimaIndex).flags.Privilegios = PlayerType.User Then
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
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                
                If UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                    N = RandomNumber(100, 1000)
                Else
                    N = RandomNumber(1, 100)
                End If
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + N
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).name & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " es un criminal!" & FONTTYPE_INFO)
    End If

    If Not Criminal(LadrOnIndex) Then
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
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= -1 Then
                    Suerte = 200
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 11 Then
                    Suerte = 190
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 21 Then
                    Suerte = 180
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 31 Then
                    Suerte = 170
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 41 Then
                    Suerte = 160
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 51 Then
                    Suerte = 150
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 61 Then
                    Suerte = 140
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 71 Then
                    Suerte = 130
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 81 Then
                    Suerte = 120
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 91 Then
                    Suerte = 110
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) = 100 Then
                    Suerte = 100
End If

If UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
    res = RandomNumber(0, Suerte)
    If res < 25 Then res = 0
Else
    res = RandomNumber(0, Suerte * 1.25)
End If

If res < 15 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 1.5)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(UserIndex).name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has apuñalado la criatura por " & Int(daño * 2) & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT)
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UCase$(UserList(UserIndex).Clase) = "LEÑADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UCase$(UserList(UserIndex).Clase) = "LEÑADOR" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Talar)

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlASALTO
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP
End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UCase$(UserList(UserIndex).Clase) = "MINERO" Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UCase$(UserList(UserIndex).Clase) = "MINERO" Then
        MiObj.Amount = RandomNumber(1, 6)
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has extraido algunos minerales!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No has conseguido nada!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Mineria)

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
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has terminado de meditar." & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
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
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has recuperado " & cant & " puntos de mana!" & FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASM" & UserList(UserIndex).Stats.MinMAN)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub



Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has logrado desarmar a tu oponente!" & FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then Call SendData(SendTarget.ToIndex, VictimIndex, 0, "||Tu oponente te ha desarmado!" & FONTTYPE_FIGHT)
    End If
End Sub

