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
      
Call WriteConsoleMsg(attackerIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Then
    EraCriminal = Criminal(attackerIndex)
    
    If (Not Criminal(VictimIndex)) Then
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
    
    If EraCriminal And Not Criminal(attackerIndex) Then
        Call RefreshCharStatus(attackerIndex)
    ElseIf Not EraCriminal And Criminal(attackerIndex) Then
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


Sub RevivirUsuario(ByVal Userindex As Integer)

UserList(Userindex).flags.Muerto = 0
UserList(Userindex).Stats.MinHP = 35

'No puede estar empollando
UserList(Userindex).flags.EstaEmpo = 0
UserList(Userindex).EmpoCont = 0

If UserList(Userindex).Stats.MinHP > UserList(Userindex).Stats.MaxHP Then
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(Userindex)
Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).OrigChar.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
Call WriteUpdateUserStats(Userindex)

End Sub

Sub ChangeUserChar(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(Userindex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
    End With
    
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChange(body, Head, heading, UserList(Userindex).Char.CharIndex, Arma, Escudo, UserList(Userindex).Char.FX, UserList(Userindex).Char.loops, casco))
End Sub

Sub EnviarFama(ByVal Userindex As Integer)
    Dim L As Long
    
    L = (-UserList(Userindex).Reputacion.AsesinoRep) + _
        (-UserList(Userindex).Reputacion.BandidoRep) + _
        UserList(Userindex).Reputacion.BurguesRep + _
        (-UserList(Userindex).Reputacion.LadronesRep) + _
        UserList(Userindex).Reputacion.NobleRep + _
        UserList(Userindex).Reputacion.PlebeRep
    L = Round(L / 6)
    
    UserList(Userindex).Reputacion.Promedio = L
    
    Call WriteFame(Userindex)
End Sub

Sub EnviarAtrib(ByVal Userindex As Integer)
    Call WriteAttributes(Userindex)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal Userindex As Integer)
    Call WriteMiniStats(Userindex)
End Sub

Sub EraseUserChar(ByVal sndIndex As Integer, ByVal Userindex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(Userindex).Char.CharIndex) = 0
    
    If UserList(Userindex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterRemove(UserList(Userindex).Char.CharIndex))
    Call QuitarUser(Userindex, UserList(Userindex).Pos.Map)
    
    MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Userindex = 0
    UserList(Userindex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Sub RefreshCharStatus(ByVal Userindex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 6/04/2007
'Refreshes the status and tag of UserIndex.
'*************************************************
    Dim klan As String
    If UserList(Userindex).guildIndex > 0 Then
        klan = modGuilds.GuildName(UserList(Userindex).guildIndex)
        klan = " <" & klan & ">"
    End If
    
    If UserList(Userindex).showName Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTag(Userindex, Criminal(Userindex), UserList(Userindex).name & klan))
    Else
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTag(Userindex, Criminal(Userindex), vbNullString))
    End If
    'Call UsUaRiOs.MakeUserChar(True, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer
    Dim userStatus As Byte

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(Userindex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(Userindex).Char.CharIndex = CharIndex
            CharList(CharIndex) = Userindex
        End If
        
        'Is the character new??
        If MapData(Map, X, Y).Userindex = Userindex Then
            userStatus = USER_VIEJO
        Else
            'Place character on map
            MapData(Map, X, Y).Userindex = Userindex
            
            userStatus = USER_NUEVO
        End If
        
        'Send make character command to clients
        Dim klan As String
        If UserList(Userindex).guildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(Userindex).guildIndex)
        End If
        
        Dim bCr As Byte
        
        bCr = Criminal(Userindex)
        
        If LenB(klan) <> 0 Then
            If Not toMap Then
                If UserList(Userindex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, UserList(Userindex).name & " <" & klan & ">", bCr, UserList(Userindex).flags.Privilegios)
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(Userindex, UserList(Userindex).Pos.Map)
                Call CheckUpdateNeededUser(Userindex, userStatus)
            End If
        Else 'if tiene clan
            If Not toMap Then
                If UserList(Userindex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, UserList(Userindex).name, bCr, UserList(Userindex).flags.Privilegios)
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(Userindex, UserList(Userindex).Pos.Map)
                Call CheckUpdateNeededUser(Userindex, userStatus)
            End If
        End If 'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(Userindex)
End Sub

Sub CheckUserLevel(ByVal Userindex As Integer)
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
'*************************************************

On Error GoTo errhandler

Dim Pts As Integer
Dim Constitucion As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim AumentoHP As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(Userindex).Stats.ELV >= STAT_MAXELV Then
    UserList(Userindex).Stats.Exp = 0
    UserList(Userindex).Stats.ELU = 0
    Exit Sub
End If
    
WasNewbie = EsNewbie(Userindex)

Do While UserList(Userindex).Stats.Exp >= UserList(Userindex).Stats.ELU
    
    'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
    'nivel
    If UserList(Userindex).Stats.ELV >= STAT_MAXELV Then
        UserList(Userindex).Stats.Exp = 0
        UserList(Userindex).Stats.ELU = 0
        Exit Sub
    End If
    
    'Store it!
    Call Statistics.UserLevelUp(Userindex)
    
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL))
    Call WriteConsoleMsg(Userindex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
    
    If UserList(Userindex).Stats.ELV = 1 Then
        Pts = 10
    Else
        'For multiple levels being rised at once
        Pts = Pts + 5
    End If
    
    UserList(Userindex).Stats.SkillPts = UserList(Userindex).Stats.SkillPts + Pts
    
    Call WriteConsoleMsg(Userindex, "Has ganado " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
       
    UserList(Userindex).Stats.ELV = UserList(Userindex).Stats.ELV + 1
    
    UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp - UserList(Userindex).Stats.ELU
    
    If Not EsNewbie(Userindex) And WasNewbie Then
        Call QuitarNewbieObj(Userindex)
        If UCase$(MapInfo(UserList(Userindex).Pos.Map).Restringir) = "NEWBIE" Then
            Call WarpUserChar(Userindex, 1, 50, 50, True)
            Call WriteConsoleMsg(Userindex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If

    'Nueva subida de exp x lvl. Pablo (ToxicWaste)
    If UserList(Userindex).Stats.ELV < 15 Then
        UserList(Userindex).Stats.ELU = UserList(Userindex).Stats.ELU * 1.4
    ElseIf UserList(Userindex).Stats.ELV < 21 Then
        UserList(Userindex).Stats.ELU = UserList(Userindex).Stats.ELU * 1.35
    ElseIf UserList(Userindex).Stats.ELV < 33 Then
        UserList(Userindex).Stats.ELU = UserList(Userindex).Stats.ELU * 1.3
    ElseIf UserList(Userindex).Stats.ELV < 41 Then
        UserList(Userindex).Stats.ELU = UserList(Userindex).Stats.ELU * 1.225
    Else
        UserList(Userindex).Stats.ELU = UserList(Userindex).Stats.ELU * 1.25
    End If
    
    Constitucion = UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion)
    
    Select Case UserList(Userindex).clase
        Case eClass.Warrior
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = IIf(UserList(Userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Hunter
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            AumentoHIT = IIf(UserList(Userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Pirat
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case eClass.Paladin
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            
            AumentoHIT = IIf(UserList(Userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Thief
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case eClass.Mage
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(5, 7)
                Case 18
                    AumentoHP = RandomNumber(4, 7)
                Case 17
                    AumentoHP = RandomNumber(4, 6)
                Case 16
                    AumentoHP = RandomNumber(3, 6)
                Case 15
                    AumentoHP = RandomNumber(3, 5)
                Case 14
                    AumentoHP = RandomNumber(2, 5)
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
            AumentoMANA = 2.8 * UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case eClass.Lumberjack
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case eClass.Miner
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case eClass.Fisher
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case eClass.Cleric
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Druid
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Assasin
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = IIf(UserList(Userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Bard
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Blacksmith, eClass.Carpenter
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
            
        Case eClass.Bandit
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = IIf(UserList(Userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = IIf(UserList(Userindex).Stats.MaxMAN = 300, 0, UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) - 10)
            If AumentoMANA < 4 Then AumentoMANA = 4
            AumentoSTA = AumentoSTLeñador
        Case Else
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, Constitucion \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(Userindex).Stats.MaxHP = UserList(Userindex).Stats.MaxHP + AumentoHP
    If UserList(Userindex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(Userindex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(Userindex).Stats.MaxSta = UserList(Userindex).Stats.MaxSta + AumentoSTA
    If UserList(Userindex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(Userindex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(Userindex).Stats.MaxMAN = UserList(Userindex).Stats.MaxMAN + AumentoMANA
    If UserList(Userindex).Stats.ELV < 36 Then
        If UserList(Userindex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(Userindex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(Userindex).Stats.MaxMAN > 9999 Then _
            UserList(Userindex).Stats.MaxMAN = 9999
    End If
    If UserList(Userindex).clase = eClass.Bandit Then 'mana del bandido restringido hasta 300
        If UserList(Userindex).Stats.MaxMAN > 300 Then
            UserList(Userindex).Stats.MaxMAN = 300
        End If
    End If
    
    'Actualizamos Golpe Máximo
    UserList(Userindex).Stats.MaxHIT = UserList(Userindex).Stats.MaxHIT + AumentoHIT
    If UserList(Userindex).Stats.ELV < 36 Then
        If UserList(Userindex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(Userindex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(Userindex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(Userindex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(Userindex).Stats.MinHIT = UserList(Userindex).Stats.MinHIT + AumentoHIT
    If UserList(Userindex).Stats.ELV < 36 Then
        If UserList(Userindex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(Userindex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(Userindex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(Userindex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then
        Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoSTA > 0 Then
        Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoMANA > 0 Then
        Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoHIT > 0 Then
        Call WriteConsoleMsg(Userindex, "Tu golpe maximo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, "Tu golpe minimo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call LogDesarrollo(Date & " " & UserList(Userindex).name & " paso a nivel " & UserList(Userindex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
Loop

'Send all gained skill points at once (if any)
If Pts > 0 Then _
    Call WriteLevelUp(Userindex, Pts)

Call WriteUpdateUserStats(Userindex)

Exit Sub

errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub

Function PuedeAtravesarAgua(ByVal Userindex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(Userindex).flags.Navegando = 1 Or _
  UserList(Userindex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal Userindex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
    
    nPos = UserList(Userindex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(Userindex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(Userindex)) Then
        If MapInfo(UserList(Userindex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendData(SendTarget.ToAllButIndex, Userindex, PrepareMessageCharacterMove(UserList(Userindex).Char.CharIndex, nPos.X, nPos.Y))

        End If
        
        'Update map and user pos
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Userindex = 0
        UserList(Userindex).Pos = nPos
        UserList(Userindex).Char.heading = nHeading
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Userindex = Userindex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(Userindex, nHeading)
    Else
        Call WritePosUpdate(Userindex)
    End If
    
    If UserList(Userindex).Counters.Trabajando Then _
        UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando - 1

    If UserList(Userindex).Counters.Ocultando Then _
        UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
    UserList(Userindex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(Userindex, Slot)
End Sub


Function NextOpenCharIndex() As Integer
'Modificada por el oso para codificar los MP1234,2,1 en 2 bytes
'para lograrlo, el charindex no puede tener su bit numero 6 (desde 0) en 1
'y tampoco puede ser un charindex que tenga el bit 0 en 1.

On Local Error GoTo hayerror

Dim LoopC As Integer
    
    LoopC = 1
    
    While LoopC < MAXCHARS
        If CharList(LoopC) = 0 And Not ((LoopC And &HFFC0&) = 64) Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        Else
            LoopC = LoopC + 1
        End If
    Wend

Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.description)

End Function

Function NextOpenUser() As Integer
    
    Dim LoopC As Integer
      
    For LoopC = 1 To MaxUsers + 1
      If LoopC > MaxUsers Then Exit For
      If (UserList(LoopC).ConnID = -1) Then Exit For
    Next LoopC
      
    NextOpenUser = LoopC

End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
Dim GuildI As Integer


    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(Userindex).Stats.ELV & "  EXP: " & UserList(Userindex).Stats.Exp & "/" & UserList(Userindex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(Userindex).Stats.MinHP & "/" & UserList(Userindex).Stats.MaxHP & "  Mana: " & UserList(Userindex).Stats.MinMAN & "/" & UserList(Userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(Userindex).Stats.MinSta & "/" & UserList(Userindex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
    If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHIT & "/" & UserList(Userindex).Stats.MaxHIT & " (" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHIT & "/" & UserList(Userindex).Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
        If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    GuildI = UserList(Userindex).guildIndex
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
        TempDate = Now - UserList(Userindex).LogOnTime
        TempSecs = (UserList(Userindex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
    #End If
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(Userindex).Stats.GLD & "  Posicion: " & UserList(Userindex).Pos.X & "," & UserList(Userindex).Pos.Y & " en mapa " & UserList(Userindex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
  
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
With UserList(Userindex)
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

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
On Error Resume Next

    Dim j As Long
    
    
    Call WriteConsoleMsg(sendIndex, UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(Userindex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(Userindex).Invent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(Userindex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(Userindex).Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)
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

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
On Error Resume Next
Dim j As Integer
Call WriteConsoleMsg(sendIndex, UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(Userindex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
Next
Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(Userindex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
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


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(Userindex).name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 24/01/2007
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'**********************************************
Dim EraCriminal As Boolean

'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(Userindex).name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(Userindex, Npclist(NpcIndex).MaestroUser)

If EsMascotaCiudadano(NpcIndex, Userindex) Then
            Call VolverCriminal(Userindex)
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    EraCriminal = Criminal(Userindex)
    
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call VolverCriminal(Userindex)
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                Call VolverCriminal(Userindex)
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlCAZADOR / 2
       If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
    If EraCriminal And Not Criminal(Userindex) Then
        Call VolverCiudadano(Userindex)
    End If
End If

End Sub

Function PuedeApuñalar(ByVal Userindex As Integer) As Boolean

If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(Userindex).clase = eClass.Assasin) And _
  (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function

Sub SubirSkill(ByVal Userindex As Integer, ByVal Skill As Integer)

If UserList(Userindex).flags.Hambre = 0 And _
  UserList(Userindex).flags.Sed = 0 Then
    
    If UserList(Userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    Dim Lvl As Integer
    Lvl = UserList(Userindex).Stats.ELV
    
    If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
    
    If UserList(Userindex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub

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
        UserList(Userindex).Stats.UserSkills(Skill) = UserList(Userindex).Stats.UserSkills(Skill) + 1
        Call WriteConsoleMsg(Userindex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(Userindex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
        
        UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + 50
        If UserList(Userindex).Stats.Exp > MAXEXP Then _
            UserList(Userindex).Stats.Exp = MAXEXP
        
        Call WriteConsoleMsg(Userindex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call WriteUpdateExp(Userindex)
        Call CheckUserLevel(Userindex)
    End If
End If

End Sub

Sub UserDie(ByVal Userindex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UserList(Userindex).genero = eGenero.Mujer Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(UserList(Userindex).Char.CharIndex))
    
    UserList(Userindex).Stats.MinHP = 0
    UserList(Userindex).Stats.MinSta = 0
    UserList(Userindex).flags.AtacadoPorNpc = 0
    UserList(Userindex).flags.AtacadoPorUser = 0
    UserList(Userindex).flags.Envenenado = 0
    UserList(Userindex).flags.Muerto = 1
    
    
    Dim aN As Integer
    
    aN = UserList(Userindex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    '<<<< Paralisis >>>>
    If UserList(Userindex).flags.Paralizado = 1 Then
        UserList(Userindex).flags.Paralizado = 0
        Call WriteParalizeOK(Userindex)
    End If
    
    '<<< Estupidez >>>
    If UserList(Userindex).flags.Estupidez = 1 Then
        UserList(Userindex).flags.Estupidez = 0
        Call WriteDumbNoMore(Userindex)
    End If
    
    '<<<< Descansando >>>>
    If UserList(Userindex).flags.Descansar Then
        UserList(Userindex).flags.Descansar = False
        Call WriteRestOK(Userindex)
    End If
    
    '<<<< Meditando >>>>
    If UserList(Userindex).flags.Meditando Then
        UserList(Userindex).flags.Meditando = False
        Call WriteMeditateToggle(Userindex)
    End If
    
    '<<<< Invisible >>>>
    If UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then
        UserList(Userindex).flags.Oculto = 0
        UserList(Userindex).Counters.TiempoOculto = 0
        UserList(Userindex).flags.invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
    End If
    
    If TriggerZonaPelea(Userindex, Userindex) <> eTrigger6.TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(Userindex) Or Criminal(Userindex) Then
            Call TirarTodo(Userindex)
        Else
            If EsNewbie(Userindex) Then Call TirarTodosLosItemsNoNewbies(Userindex)
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(Userindex).Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.AnilloEqpSlot)
    End If
    'desequipar municiones
    If UserList(Userindex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(Userindex, UserList(Userindex).Invent.EscudoEqpSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(Userindex).Char.loops = LoopAdEternum Then
        UserList(Userindex).Char.FX = 0
        UserList(Userindex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(Userindex).flags.Mimetizado = 1 Then
        UserList(Userindex).Char.body = UserList(Userindex).CharMimetizado.body
        UserList(Userindex).Char.Head = UserList(Userindex).CharMimetizado.Head
        UserList(Userindex).Char.CascoAnim = UserList(Userindex).CharMimetizado.CascoAnim
        UserList(Userindex).Char.ShieldAnim = UserList(Userindex).CharMimetizado.ShieldAnim
        UserList(Userindex).Char.WeaponAnim = UserList(Userindex).CharMimetizado.WeaponAnim
        UserList(Userindex).Counters.Mimetismo = 0
        UserList(Userindex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(Userindex).flags.Navegando = 0 Then
        UserList(Userindex).Char.body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
    Else
        UserList(Userindex).Char.body = iFragataFantasmal ';)
    End If
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(Userindex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(Userindex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(Userindex).MascotasIndex(i)).Movement = Npclist(UserList(Userindex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(Userindex).MascotasIndex(i)).Hostile = Npclist(UserList(Userindex).MascotasIndex(i)).flags.OldHostil
                    UserList(Userindex).MascotasIndex(i) = 0
                    UserList(Userindex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(Userindex).NroMacotas = 0
    
    
    'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    '        Dim MiObj As Obj
    '        Dim nPos As WorldPos
    '        MiObj.ObjIndex = RandomNumber(554, 555)
    '        MiObj.Amount = 1
    '        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    '        Dim ManchaSangre As New cGarbage
    '        ManchaSangre.Map = nPos.Map
    '        ManchaSangre.X = nPos.X
    '        ManchaSangre.Y = nPos.Y
    '        Call TrashCollector.Add(ManchaSangre)
    'End If
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(Userindex)
    
    
    '<<Castigos por party>>
    If UserList(Userindex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(Userindex, UserList(Userindex).Stats.ELV * -10 * mdParty.CantMiembros(Userindex), UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
    End If

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If Criminal(Muerto) Then
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
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + Obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
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

Sub WarpUserChar(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    'Quitar el dialogo
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(UserList(Userindex).Char.CharIndex))
    
    Call WriteRemoveAllDialogs(Userindex)
    
    OldMap = UserList(Userindex).Pos.Map
    OldX = UserList(Userindex).Pos.X
    OldY = UserList(Userindex).Pos.Y
    
    Call EraseUserChar(OldMap, Userindex)
    
    If OldMap <> Map Then
        Call WriteChangeMap(Userindex, Map, MapInfo(UserList(Userindex).Pos.Map).MapVersion)
        Call WritePlayMidi(Userindex, ReadField(1, MapInfo(Map).Music, 45))
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(Userindex).Pos.X = X
    UserList(Userindex).Pos.Y = Y
    UserList(Userindex).Pos.Map = Map
    
    Call MakeUserChar(True, Map, Userindex, Map, X, Y)
    Call WriteUserCharIndexInServer(Userindex)
    
    'Force a flush, so user index is in there before it's destroyed for teleporting
    Call FlushBuffer(Userindex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1) And (Not UserList(Userindex).flags.AdminInvisible = 1) Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))
    End If
    
    If FX And UserList(Userindex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_WARP))
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    Call WarpMascotas(Userindex)
End Sub

Sub WarpMascotas(ByVal Userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(Userindex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
                UserList(Userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(Userindex, "Pierdes el control de tus mascotas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(Userindex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(Userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(Userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(Userindex).Pos, False, PetRespawn(i))
            UserList(Userindex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(Userindex).MascotasIndex(i) = 0 Then
                UserList(Userindex).MascotasIndex(i) = 0
                UserList(Userindex).MascotasType(i) = 0
                If UserList(Userindex).NroMacotas > 0 Then UserList(Userindex).NroMacotas = UserList(Userindex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = Userindex
            Npclist(UserList(Userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(Userindex).MascotasIndex(i)).Target = 0
            Npclist(UserList(Userindex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(Userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(Userindex).MascotasIndex(i))
        End If
    Next i
    
    UserList(Userindex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal Userindex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(Userindex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(Userindex).NroMacotas Then UserList(Userindex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal Userindex As Integer, Optional ByVal Tiempo As Integer = -1)
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(Userindex).flags.UserLogged And Not UserList(Userindex).Counters.Saliendo Then
        UserList(Userindex).Counters.Saliendo = True
        UserList(Userindex).Counters.Salir = IIf((UserList(Userindex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(Userindex).Pos.Map).Pk, Tiempo, 0)
        
        Call WriteConsoleMsg(Userindex, "Cerrando...Se cerrará el juego en " & UserList(Userindex).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal Userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
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

Public Sub Empollando(ByVal Userindex As Integer)
If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    UserList(Userindex).flags.EstaEmpo = 1
Else
    UserList(Userindex).flags.EstaEmpo = 0
    UserList(Userindex).EmpoCont = 0
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
Exit Sub

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

Sub VolverCriminal(ByVal Userindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente
'**************************************************************
If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

If UserList(Userindex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
    UserList(Userindex).Reputacion.BurguesRep = 0
    UserList(Userindex).Reputacion.NobleRep = 0
    UserList(Userindex).Reputacion.PlebeRep = 0
    UserList(Userindex).Reputacion.BandidoRep = UserList(Userindex).Reputacion.BandidoRep + vlASALTO
    If UserList(Userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(Userindex).Reputacion.BandidoRep = MAXREP
    If UserList(Userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)
End If

Call RefreshCharStatus(Userindex)

End Sub

Sub VolverCiudadano(ByVal Userindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************

If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

UserList(Userindex).Reputacion.LadronesRep = 0
UserList(Userindex).Reputacion.BandidoRep = 0
UserList(Userindex).Reputacion.AsesinoRep = 0
UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlASALTO
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

Call RefreshCharStatus(Userindex)

End Sub

