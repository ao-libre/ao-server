Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)

Dim DaExp As Integer
Dim EraCriminal As Boolean

DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp
If UserList(attackerIndex).Stats.Exp > MAXEXP Then _
    UserList(attackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
      
Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Then
    EraCriminal = criminal(attackerIndex)
    
    If (Not criminal(VictimIndex)) Then
         UserList(attackerIndex).Reputacion.AsesinoRep = UserList(attackerIndex).Reputacion.AsesinoRep + vlASESINO * 2
         If UserList(attackerIndex).Reputacion.AsesinoRep > MAXREP Then _
            UserList(attackerIndex).Reputacion.AsesinoRep = MAXREP
         UserList(attackerIndex).Reputacion.BurguesRep = 0
         UserList(attackerIndex).Reputacion.NobleRep = 0
         UserList(attackerIndex).Reputacion.PlebeRep = 0
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
End If

Call UserDie(VictimIndex)

If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then _
    UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1

Call FlushBuffer(VictimIndex)

'Log
Call LogAsesinato(UserList(attackerIndex).name & " asesino a " & UserList(VictimIndex).name)

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)

'No puede estar empollando
UserList(UserIndex).flags.EstaEmpo = 0
UserList(UserIndex).EmpoCont = 0

If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
End If

If UserList(UserIndex).flags.Navegando = 1 Then
    Dim Barco As ObjData
    Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    UserList(UserIndex).Char.Head = 0
    
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
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
    Call DarCuerpoDesnudo(UserIndex)
    
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
End If

Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteUpdateUserStats(UserIndex)

End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
    End With
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, casco))
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim L As Long
    
    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
        (-UserList(UserIndex).Reputacion.BandidoRep) + _
        UserList(UserIndex).Reputacion.BurguesRep + _
        (-UserList(UserIndex).Reputacion.LadronesRep) + _
        UserList(UserIndex).Reputacion.NobleRep + _
        UserList(UserIndex).Reputacion.PlebeRep
    L = Round(L / 6)
    
    UserList(UserIndex).Reputacion.Promedio = L
    
    Call WriteFame(UserIndex)
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex))
    Call QuitarUser(UserIndex, UserList(UserIndex).Pos.map)
    
    MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/21/2008 (NicoNZ)
'Refreshes the status and tag of UserIndex.
'*************************************************
    Dim klan As String
    Dim Barco As ObjData
    
    If UserList(UserIndex).guildIndex > 0 Then
        klan = modGuilds.GuildName(UserList(UserIndex).guildIndex)
        klan = " <" & klan & ">"
    End If
    
    If UserList(UserIndex).showName Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, criminal(UserIndex), UserList(UserIndex).name & klan))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, criminal(UserIndex), vbNullString))
    End If
    
    'Si esta navengando, se cambia la barca.
    If UserList(UserIndex).flags.Navegando Then
        Barco = ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex)
        
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
        
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    End If
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(map, X, Y) Then
        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If
        
        'Place character on map if needed
        If toMap Then _
            MapData(map, X, Y).UserIndex = UserIndex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(UserIndex).guildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(UserIndex).guildIndex)
        End If
        
        Dim bCr As Byte
        
        bCr = criminal(UserIndex)
        
        If LenB(klan) <> 0 Then
            If Not toMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name & " <" & klan & ">", bCr, UserList(UserIndex).flags.Privilegios)
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        Else 'if tiene clan
            If Not toMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name, bCr, UserList(UserIndex).flags.Privilegios)
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        End If 'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 01/10/2007
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
'*************************************************

On Error GoTo Errhandler

Dim Pts As Integer
Dim Constitucion As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim AumentoHP As Integer
Dim WasNewbie As Boolean
Dim Promedio As Double
Dim aux As Integer
Dim DistVida(1 To 5) As Integer

'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If
    
WasNewbie = EsNewbie(UserIndex)

Do While UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU
    
    'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo nivel
    If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
        UserList(UserIndex).Stats.Exp = 0
        UserList(UserIndex).Stats.ELU = 0
        Exit Sub
    End If
    
    'Store it!
    Call Statistics.UserLevelUp(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call WriteConsoleMsg(UserIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
    
    If UserList(UserIndex).Stats.ELV = 1 Then
        Pts = 10
    Else
        'For multiple levels being rised at once
        Pts = Pts + 5
    End If
    
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
    
    'Nueva subida de exp x lvl. Pablo (ToxicWaste)
    If UserList(UserIndex).Stats.ELV < 15 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.4
    ElseIf UserList(UserIndex).Stats.ELV < 21 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.35
    ElseIf UserList(UserIndex).Stats.ELV < 33 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    ElseIf UserList(UserIndex).Stats.ELV < 41 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.225
    Else
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.25
    End If
    
    Constitucion = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
    
    'Calculo subida de vida
    Promedio = ModVida(UserList(UserIndex).clase) - (21 - Constitucion) * 0.5
    aux = RandomNumber(0, 100)
    

    If Promedio - Int(Promedio) = 0.5 Then
    'Es promedio semientero
        DistVida(1) = DistribucionSemienteraVida(1)
        DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
        DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
        DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
        
        If aux <= DistVida(1) Then
            AumentoHP = Promedio + 1.5
        ElseIf aux <= DistVida(2) Then
            AumentoHP = Promedio + 0.5
        ElseIf aux <= DistVida(3) Then
            AumentoHP = Promedio - 0.5
        Else
            AumentoHP = Promedio - 1.5
        End If
    Else
    'Es promedio entero
        'Sacar este IF en la 0.13 (ToxicWaste)
        If UserList(UserIndex).clase = eClass.Mage Then
        
        If aux <= 33 Then
            AumentoHP = Promedio + 1
        ElseIf aux <= 66 Then
            AumentoHP = Promedio
        Else
            AumentoHP = Promedio - 1
        End If
        
        Else
            DistVida(1) = DistribucionSemienteraVida(1)
            DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
            DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
            DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
            DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
            
            If aux <= DistVida(1) Then
                AumentoHP = Promedio + 2
            ElseIf aux <= DistVida(2) Then
                AumentoHP = Promedio + 1
            ElseIf aux <= DistVida(3) Then
                AumentoHP = Promedio
            ElseIf aux <= DistVida(4) Then
                AumentoHP = Promedio - 1
            Else
                AumentoHP = Promedio - 2
            End If
        End If
    End If

    Select Case UserList(UserIndex).clase
        Case eClass.Warrior
        
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Hunter

            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Pirat

            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case eClass.Paladin
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Thief

            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case eClass.Mage
            
            AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
            AumentoMANA = 2.8 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case eClass.Lumberjack
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case eClass.Miner
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case eClass.Fisher
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case eClass.Cleric
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Druid
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Assasin
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Bard
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Blacksmith, eClass.Carpenter
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
            
        Case eClass.Bandit
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = IIf(UserList(UserIndex).Stats.MaxMAN = 300, 0, UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 10)
            If AumentoMANA < 4 Then AumentoMANA = 4
            AumentoSTA = AumentoSTLeñador
        Case Else
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + AumentoHP
    If UserList(UserIndex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(UserIndex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(UserIndex).Stats.MaxSta = UserList(UserIndex).Stats.MaxSta + AumentoSTA
    If UserList(UserIndex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(UserIndex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + AumentoMANA
    If UserList(UserIndex).Stats.ELV < 36 Then
        If UserList(UserIndex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(UserIndex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(UserIndex).Stats.MaxMAN > 9999 Then _
            UserList(UserIndex).Stats.MaxMAN = 9999
    End If
    If UserList(UserIndex).clase = eClass.Bandit Then 'mana del bandido restringido hasta 300
        If UserList(UserIndex).Stats.MaxMAN > 300 Then
            UserList(UserIndex).Stats.MaxMAN = 300
        End If
    End If
    
    'Actualizamos Golpe Máximo
    UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT + AumentoHIT
    If UserList(UserIndex).Stats.ELV < 36 Then
        If UserList(UserIndex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(UserIndex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(UserIndex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(UserIndex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + AumentoHIT
    If UserList(UserIndex).Stats.ELV < 36 Then
        If UserList(UserIndex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(UserIndex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(UserIndex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(UserIndex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then
        Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoSTA > 0 Then
        Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoMANA > 0 Then
        Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoHIT > 0 Then
        Call WriteConsoleMsg(UserIndex, "Tu golpe maximo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Tu golpe minimo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call LogDesarrollo(UserList(UserIndex).name & " paso a nivel " & UserList(UserIndex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
Loop

'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
If Not EsNewbie(UserIndex) And WasNewbie Then
    Call QuitarNewbieObj(UserIndex)
    If UCase$(MapInfo(UserList(UserIndex).Pos.map).Restringir) = "NEWBIE" Then
        Call WarpUserChar(UserIndex, 1, 50, 50, True)
        Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

'Send all gained skill points at once (if any)
If Pts > 0 Then
    Call WriteLevelUp(UserIndex, Pts)
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
End If

Call WriteUpdateUserStats(UserIndex)

Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

Dim nPos As WorldPos
Dim sailing As Boolean


    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(UserIndex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))

        End If
        
        'Update map and user pos
        MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.heading = nHeading
        MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
Dim GuildI As Integer


    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    GuildI = UserList(UserIndex).guildIndex
    If GuildI > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
            Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
    End If
    
    #If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - UserList(UserIndex).LogOnTime
        TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
    #End If
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.map, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
  
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
With UserList(UserIndex)
    Call WriteConsoleMsg(sendIndex, "Pj: " & .name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
    If .Faccion.ArmadaReal = 1 Then
        Call WriteConsoleMsg(sendIndex, "Armada Real Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
    ElseIf .Faccion.FuerzasCaos = 1 Then
        Call WriteConsoleMsg(sendIndex, "Legion Oscura Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
    ElseIf .Faccion.RecibioExpInicialReal = 1 Then
        Call WriteConsoleMsg(sendIndex, "Fue Armada Real", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
    ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
        Call WriteConsoleMsg(sendIndex, "Fue Legionario", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
    End If
    Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
    If .guildIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.guildIndex), FontTypeNames.FONTTYPE_INFO)
    End If
    
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Armada Real Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion Oscura Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Armada Real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Legionario", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
Next
Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
End If

End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 06/28/2008
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'**********************************************
Dim EraCriminal As Boolean

'Guardamos el usuario que ataco el npc.
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).name

'Npc que estabas atacando.
Dim LastNpcHit As Integer
LastNpcHit = UserList(UserIndex).flags.NPCAtacado
'Guarda el NPC que estas atacando ahora.
UserList(UserIndex).flags.NPCAtacado = NpcIndex

'Revisamos robo de npc.
'Guarda el primer nick que lo ataca.
If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
    'El que le pegabas antes ya no es tuyo
    If LastNpcHit <> 0 Then
        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
        End If
    End If
    Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).name
ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).name Then
    'Estas robando NPC
    'El que le pegabas antes ya no es tuyo
    If LastNpcHit <> 0 Then
        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
        End If
    End If
End If

If Npclist(NpcIndex).MaestroUser > 0 Then
    If Npclist(NpcIndex).MaestroUser <> UserIndex Then
        Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
    End If
End If

If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    EraCriminal = criminal(UserIndex)
    
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call VolverCriminal(UserIndex)
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                Call VolverCriminal(UserIndex)
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP
    End If
    
    If Npclist(NpcIndex).MaestroUser <> UserIndex Then
        'hacemos que el npc se defienda
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    End If
    
    If EraCriminal And Not criminal(UserIndex) Then
        Call VolverCiudadano(UserIndex)
    End If
End If

End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).clase = eClass.Assasin) And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

    If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        
        If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
        
        Dim Lvl As Integer
        Lvl = UserList(UserIndex).Stats.ELV
        
        If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
        
        If UserList(UserIndex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
    
        Dim Aumenta As Integer
        Dim Prob As Integer
        
        If Lvl <= 3 Then
            Prob = 25
        ElseIf Lvl > 3 And Lvl < 6 Then
            Prob = 35
        ElseIf Lvl >= 6 And Lvl < 10 Then
            Prob = 40
        ElseIf Lvl >= 10 And Lvl < 20 Then
            Prob = 45
        Else
            Prob = 50
        End If
        
        Aumenta = RandomNumber(1, Prob)
        
        If Aumenta = 7 Then
            UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
            Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + 50
            If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                UserList(UserIndex).Stats.Exp = MAXEXP
            
            Call WriteConsoleMsg(UserIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
            
            Call WriteUpdateExp(UserIndex)
            Call CheckUserLevel(UserIndex)
        End If
    End If

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 04/15/2008 (NicoNZ)
'Ahora se resetea el counter del invi
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    'Sonido
    If UserList(UserIndex).genero = eGenero.Mujer Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1
    UserList(UserIndex).flags.SeguroResu = True
    Call WriteResuscitationSafeOn(UserIndex)
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    aN = UserList(UserIndex).flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)
    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call WriteRestOK(UserIndex)
    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
    End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
    
    If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(UserIndex) Or criminal(UserIndex) Then
            Call TirarTodo(UserIndex)
        Else
            If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
    End If
    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).Char.loops = INFINITE_LOOPS Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
        UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
    End If
    
    ' << Restauramos los atributos >>
    If UserList(UserIndex).flags.TomoPocion = True Then
        For i = 1 To 5
            UserList(UserIndex).Stats.UserAtributos(i) = UserList(UserIndex).Stats.UserAtributosBackUP(i)
        Next i
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).Char.body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).Char.body = iFragataFantasmal ';)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    Next i
    
    UserList(UserIndex).NroMascotas = 0
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(UserIndex)
    
    
    '<<Castigos por party>>
    If UserList(UserIndex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(UserIndex, UserList(UserIndex).Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    End If

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
        
        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
            'con esto evitamos que se vuelva a reenlistar
        End If
    Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
        End If
    End If
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.map = Pos.map
    
    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.amount + Obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    'Quitar el dialogo
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    Call WriteRemoveAllDialogs(UserIndex)
    
    OldMap = UserList(UserIndex).Pos.map
    OldX = UserList(UserIndex).Pos.X
    OldY = UserList(UserIndex).Pos.Y
    
    Call EraseUserChar(UserIndex)
    
    If OldMap <> map Then
        Call WriteChangeMap(UserIndex, map, MapInfo(UserList(UserIndex).Pos.map).MapVersion)
        Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(map).Music, 45)))
        
        'Update new Map Users
        MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
        
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.map = map
    
    Call MakeUserChar(True, map, UserIndex, map, X, Y)
    Call WriteUserCharIndexInServer(UserIndex)
    
    'Force a flush, so user index is in there before it's destroyed for teleporting
    Call FlushBuffer(UserIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
    End If
    
    If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    Call WarpMascotas(UserIndex)
End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(UserIndex).NroMascotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            PetRespawn(i) = True
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = 0
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
          If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
            'Controlamos que se sumoneo OK
            If UserList(UserIndex).MascotasIndex(i) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
                Exit For
            End If
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
          Else
            Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
            Exit For
          End If
        End If
    Next i
    
    UserList(UserIndex).NroMascotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(UserIndex).NroMascotas Then UserList(UserIndex).NroMascotas = 0

End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 06/28/08 (NicoNZ)
'
'***************************************************
    Dim isNotVisible As Boolean
    
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        
        isNotVisible = (UserList(UserIndex).flags.Oculto Or UserList(UserIndex).flags.invisible)
        If isNotVisible Then
            UserList(UserIndex).flags.Oculto = 0
            UserList(UserIndex).flags.invisible = 0
            UserList(UserIndex).Counters.Invisibilidad = 0
            UserList(UserIndex).Counters.TiempoOculto = 0
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
        End If
        
        Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub

Public Sub Empollando(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
    Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
    
    Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
    
#If ConUpTime Then
    Dim TempSecs As Long
    Dim TempStr As String
    TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
    TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
    Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If

End If

End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & charName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente
'**************************************************************
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If

Call RefreshCharStatus(UserIndex)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlASALTO
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

Call RefreshCharStatus(UserIndex)

End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function
