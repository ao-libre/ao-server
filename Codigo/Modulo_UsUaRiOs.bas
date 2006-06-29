Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp
If UserList(AttackerIndex).Stats.Exp > MAXEXP Then _
    UserList(AttackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Has matado a " & UserList(VictimIndex).name & "!" & FONTTYPE_FIGHT)
Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT)
      
Call SendData(SendTarget.ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).name & " te ha matado!" & FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
    If (Not criminal(VictimIndex)) Then
         UserList(AttackerIndex).Reputacion.AsesinoRep = UserList(AttackerIndex).Reputacion.AsesinoRep + vlASESINO * 2
         If UserList(AttackerIndex).Reputacion.AsesinoRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.AsesinoRep = MAXREP
         UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
    Else
         UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble
         If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.NobleRep = MAXREP
    End If
End If

Call UserDie(VictimIndex)

If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then _
    UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1

'Log
Call LogAsesinato(UserList(AttackerIndex).name & " asesino a " & UserList(VictimIndex).name)

End Sub


Sub RevivirUsuario(ByVal userindex As Integer)

UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = 35

'No puede estar empollando
UserList(userindex).flags.EstaEmpo = 0
UserList(userindex).EmpoCont = 0

If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(userindex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendUserStatsBox(userindex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, _
                    ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    UserList(userindex).Char.body = body
    UserList(userindex).Char.Head = Head
    UserList(userindex).Char.heading = heading
    UserList(userindex).Char.WeaponAnim = Arma
    UserList(userindex).Char.ShieldAnim = Escudo
    UserList(userindex).Char.CascoAnim = casco
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "CP" & UserList(userindex).Char.CharIndex & "," & body & "," & Head & "," & heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).Char.CharIndex & "," & body & "," & Head & "," & heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & casco)
    End If
End Sub

Sub EnviarSubirNivel(ByVal userindex As Integer, ByVal Puntos As Integer)
    Call SendData(SendTarget.ToIndex, userindex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal userindex As Integer)
    Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
       cad = cad & UserList(userindex).Stats.UserSkills(i) & ","
    Next i
    
    SendData SendTarget.ToIndex, userindex, 0, "SKILLS" & cad$
End Sub

Sub EnviarFama(ByVal userindex As Integer)
    Dim cad As String
    
    cad = cad & UserList(userindex).Reputacion.AsesinoRep & ","
    cad = cad & UserList(userindex).Reputacion.BandidoRep & ","
    cad = cad & UserList(userindex).Reputacion.BurguesRep & ","
    cad = cad & UserList(userindex).Reputacion.LadronesRep & ","
    cad = cad & UserList(userindex).Reputacion.NobleRep & ","
    cad = cad & UserList(userindex).Reputacion.PlebeRep & ","
    
    Dim L As Long
    
    L = (-UserList(userindex).Reputacion.AsesinoRep) + _
        (-UserList(userindex).Reputacion.BandidoRep) + _
        UserList(userindex).Reputacion.BurguesRep + _
        (-UserList(userindex).Reputacion.LadronesRep) + _
        UserList(userindex).Reputacion.NobleRep + _
        UserList(userindex).Reputacion.PlebeRep
    L = Round(L / 6)
    
    UserList(userindex).Reputacion.Promedio = L
    
    cad = cad & UserList(userindex).Reputacion.Promedio
    
    SendData SendTarget.ToIndex, userindex, 0, "FAMA" & cad
End Sub

Sub EnviarAtrib(ByVal userindex As Integer)
Dim i As Integer
Dim cad As String
For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(userindex).Stats.UserAtributos(i) & ","
Next
Call SendData(SendTarget.ToIndex, userindex, 0, "ATR" & cad)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal userindex As Integer)
With UserList(userindex)
    Call SendData(SendTarget.ToIndex, userindex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & _
                .Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
                .Stats.NPCsMuertos & "," & .clase & "," & .Counters.Pena)
End With

End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(userindex).Char.CharIndex) = 0
    
    If UserList(userindex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "BP" & UserList(userindex).Char.CharIndex)
        Call QuitarUser(userindex, UserList(userindex).Pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BP" & UserList(userindex).Char.CharIndex)
    End If
    
    MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = 0
    UserList(userindex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(userindex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(userindex).Char.CharIndex = CharIndex
            CharList(CharIndex) = userindex
        End If
        
        'Place character on map
        MapData(Map, X, Y).userindex = userindex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(UserIndex).GuildIndex > 0 Then
            klan = modGuilds.guildName(UserList(UserIndex).GuildIndex)
        End If
        
        Dim bCr As Byte
        Dim SendPrivilegios As Byte
       
        bCr = criminal(userindex)

        If klan <> "" Then
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        Else
                            'Hide the name and clan
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        Else
                            'Hide the name and clan
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        End If
                    Else
                        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).Pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
        Else 'if tiene clan
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        Else
                            'Hide the name
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        Else
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        End If
                    Else
                        Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(userindex).Char.body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & bCr & "," & IIf(UserList(userindex).flags.PertAlCons = 1, 4, IIf(UserList(userindex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).Pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
       End If   'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(userindex)
End Sub

Sub CheckUserLevel(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 05/27/2006
'Checkea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'05/27/2006 Integer - Uso del switch para mejor performance y claridad.
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
If UserList(userindex).Stats.ELV = STAT_MAXELV Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(userindex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU
    
    'Store it!
    Call Statistics.UserLevelUp(userindex)
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_NIVEL)
    Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    If UserList(userindex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 5
    End If
    
    UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts + Pts
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
       
    UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
    
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.ELU
    
    If Not EsNewbie(userindex) And WasNewbie Then
        Call QuitarNewbieObj(userindex)
        If UCase$(MapInfo(UserList(userindex).Pos.Map).Restringir) = "SI" Then
            Call WarpUserChar(userindex, 1, 50, 50, True)
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes abandonar el Dungeon Newbie." & FONTTYPE_WARNING)
        End If
    End If

    If UserList(userindex).Stats.ELV < 11 Then
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.5
    ElseIf UserList(userindex).Stats.ELV < 25 Then
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.3
    Else
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.2
    End If
    
    Constitucion = UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)

    Select Case UCase$(UserList(userindex).clase)
        Case "GUERRERO"
            AumentoHP = RandomNumber(Switch(Constitucion >= 20, 8 _
                                            , Constitucion >= 18, 7 _
                                            , Constitucion >= 16, 6 _
                                            , Constitucion >= 14, 5 _
                                            , True, 4), _
                                     Switch(Constitucion = 21, 12 _
                                            , Constitucion >= 19, 11 _
                                            , Constitucion >= 17, 10 _
                                            , Constitucion >= 15, 9 _
                                            , Constitucion >= 13, 8 _
                                            , True, 7))
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 8 _
                                            , Constitucion >= 19, 7 _
                                            , Constitucion >= 17, 6 _
                                            , Constitucion >= 15, 5 _
                                            , Constitucion >= 13, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion >= 20, 11 _
                                            , Constitucion >= 18, 10 _
                                            , Constitucion >= 16, 9 _
                                            , Constitucion >= 14, 8 _
                                            , True, 7))
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PIRATA"
            AumentoHP = RandomNumber(Switch(Constitucion >= 20, 8 _
                                            , Constitucion >= 18, 7 _
                                            , Constitucion >= 16, 6 _
                                            , Constitucion >= 14, 5 _
                                            , Constitucion >= 12, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion = 21, 12 _
                                            , Constitucion >= 19, 11 _
                                            , Constitucion >= 17, 10 _
                                            , Constitucion >= 15, 9 _
                                            , Constitucion >= 13, 8 _
                                            , True, 7))
            
            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            AumentoHP = RandomNumber(Switch(Constitucion >= 20, 8 _
                                            , Constitucion >= 18, 7 _
                                            , Constitucion >= 16, 6 _
                                            , Constitucion >= 14, 5 _
                                            , Constitucion >= 12, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion = 21, 12 _
                                            , Constitucion >= 19, 11 _
                                            , Constitucion >= 17, 10 _
                                            , Constitucion >= 15, 9 _
                                            , Constitucion >= 13, 8 _
                                            , True, 7))
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "LADRON"
            AumentoHP = RandomNumber(Switch(Constitucion >= 20, 8 _
                                            , Constitucion >= 18, 7 _
                                            , Constitucion >= 16, 6 _
                                            , Constitucion >= 14, 5 _
                                            , Constitucion >= 12, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion = 21, 12 _
                                            , Constitucion >= 19, 11 _
                                            , Constitucion >= 17, 10 _
                                            , Constitucion >= 15, 9 _
                                            , Constitucion >= 13, 8 _
                                            , True, 7))
            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case "MAGO"
            AumentoHP = RandomNumber(Switch(Constitucion = 20, 6 _
                                            , Constitucion >= 18, 5 _
                                            , Constitucion >= 16, 3 _
                                            , Constitucion >= 14, 2), _
                                     Switch(Constitucion >= 20, 8 _
                                            , Constitucion >= 17, 7 _
                                            , Constitucion >= 15, 6 _
                                            , True, 5))
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case "LEÑADOR"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 8 _
                                            , Constitucion >= 19, 7 _
                                            , Constitucion >= 17, 6 _
                                            , Constitucion >= 15, 5 _
                                            , Constitucion >= 13, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion >= 20, 11 _
                                            , Constitucion >= 18, 10 _
                                            , Constitucion >= 16, 9 _
                                            , Constitucion >= 14, 8 _
                                            , True, 7))
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case "MINERO"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 8 _
                                            , Constitucion >= 19, 7 _
                                            , Constitucion >= 17, 6 _
                                            , Constitucion >= 15, 5 _
                                            , Constitucion >= 13, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion >= 20, 11 _
                                            , Constitucion >= 18, 10 _
                                            , Constitucion >= 16, 9 _
                                            , Constitucion >= 14, 8 _
                                            , True, 7))
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case "PESCADOR"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 8 _
                                            , Constitucion >= 19, 7 _
                                            , Constitucion >= 17, 6 _
                                            , Constitucion >= 15, 5 _
                                            , Constitucion >= 13, 4 _
                                            , Constitucion = 12, 3), _
                                     Switch(Constitucion >= 20, 11 _
                                            , Constitucion >= 18, 10 _
                                            , Constitucion >= 16, 9 _
                                            , Constitucion >= 14, 8 _
                                            , True, 7))
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case "CLERIGO"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 7 _
                                            , Constitucion >= 19, 6 _
                                            , Constitucion >= 17, 5 _
                                            , Constitucion >= 15, 4 _
                                            , Constitucion >= 13, 3 _
                                            , Constitucion = 12, 2), _
                                     Switch(Constitucion >= 20, 10 _
                                            , Constitucion >= 18, 9 _
                                            , Constitucion >= 16, 8 _
                                            , Constitucion >= 14, 7 _
                                            , True, 6))
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "DRUIDA"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 7 _
                                            , Constitucion >= 19, 6 _
                                            , Constitucion >= 17, 5 _
                                            , Constitucion >= 15, 4 _
                                            , Constitucion >= 13, 3 _
                                            , Constitucion = 12, 2), _
                                     Switch(Constitucion >= 20, 10 _
                                            , Constitucion >= 18, 9 _
                                            , Constitucion >= 16, 8 _
                                            , Constitucion >= 14, 7 _
                                            , True, 6))
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 7 _
                                            , Constitucion >= 19, 6 _
                                            , Constitucion >= 17, 5 _
                                            , Constitucion >= 15, 4 _
                                            , Constitucion >= 13, 3 _
                                            , Constitucion = 12, 2), _
                                     Switch(Constitucion >= 20, 10 _
                                            , Constitucion >= 18, 9 _
                                            , Constitucion >= 16, 8 _
                                            , Constitucion >= 14, 7 _
                                            , True, 6))
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 7 _
                                            , Constitucion >= 19, 6 _
                                            , Constitucion >= 17, 5 _
                                            , Constitucion >= 15, 4 _
                                            , Constitucion >= 13, 3 _
                                            , Constitucion = 12, 2), _
                                     Switch(Constitucion >= 20, 10 _
                                            , Constitucion >= 18, 9 _
                                            , Constitucion >= 16, 8 _
                                            , Constitucion >= 14, 7 _
                                            , True, 6))
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case Else
            AumentoHP = RandomNumber(Switch(Constitucion = 21, 7 _
                                            , Constitucion >= 19, 6 _
                                            , Constitucion >= 17, 5 _
                                            , Constitucion >= 15, 4 _
                                            , Constitucion >= 13, 3 _
                                            , Constitucion = 12, 2), _
                                     Switch(Constitucion >= 20, 10 _
                                            , Constitucion >= 18, 9 _
                                            , Constitucion >= 16, 8 _
                                            , Constitucion >= 14, 7 _
                                            , True, 6))

            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + AumentoHP
    If UserList(userindex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(userindex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MaxSta + AumentoSTA
    If UserList(userindex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(userindex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(userindex).Stats.MaxMAN = UserList(userindex).Stats.MaxMAN + AumentoMANA
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(userindex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(userindex).Stats.MaxMAN > 9999 Then _
            UserList(userindex).Stats.MaxMAN = 9999
    End If
    
    'Actualizamos Golpe Máximo
    UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then SendData SendTarget.ToIndex, userindex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoSTA > 0 Then SendData SendTarget.ToIndex, userindex, 0, "||Has ganado " & AumentoSTA & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData SendTarget.ToIndex, userindex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData SendTarget.ToIndex, userindex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData SendTarget.ToIndex, userindex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    Call LogDesarrollo(Date & " " & UserList(userindex).name & " paso a nivel " & UserList(userindex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, Pts)
   
    SendUserStatsBox userindex
    
Loop
'End If


Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(userindex).flags.Navegando = 1 Or _
  UserList(userindex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
    
    nPos = UserList(userindex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(userindex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(userindex)) Then
        If MapInfo(UserList(userindex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...
#If SeguridadAlkon Then
            Call SendCryptedMoveChar(nPos.Map, userindex, nPos.X, nPos.Y)
#Else
            Call SendToUserAreaButindex(userindex, "+" & UserList(userindex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
#End If
        End If
        
        'Update map and user pos
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = 0
        UserList(userindex).Pos = nPos
        UserList(userindex).Char.heading = nHeading
        MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = userindex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
    End If
    
    If UserList(userindex).Counters.Trabajando Then _
        UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

    If UserList(userindex).Counters.Ocultando Then _
        UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(userindex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(userindex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ 3)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")
    End If

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
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.Description)

End Function

Function NextOpenUser() As Integer
    
    Dim LoopC As Integer
      
    For LoopC = 1 To MaxUsers + 1
      If LoopC > MaxUsers Then Exit For
      If (UserList(LoopC).ConnID = -1) Then Exit For
    Next LoopC
      
    NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(ByVal userindex As Integer)
    Call SendData(SendTarget.ToIndex, userindex, 0, "EST" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta & "," & UserList(userindex).Stats.GLD & "," & UserList(userindex).Stats.ELV & "," & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp)
End Sub

Sub EnviarHambreYsed(ByVal userindex As Integer)
    Call SendData(SendTarget.ToIndex, userindex, 0, "EHYS" & UserList(userindex).Stats.MaxAGU & "," & UserList(userindex).Stats.MinAGU & "," & UserList(userindex).Stats.MaxHam & "," & UserList(userindex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Nivel: " & UserList(userindex).Stats.ELV & "  EXP: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Salud: " & UserList(userindex).Stats.MinHP & "/" & UserList(userindex).Stats.MaxHP & "  Mana: " & UserList(userindex).Stats.MinMAN & "/" & UserList(userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(userindex).Stats.MinSta & "/" & UserList(userindex).Stats.MaxSta & FONTTYPE_INFO)
    
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & " (" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    GuildI = UserList(userindex).GuildIndex
    If GuildI > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clan: " & modGuilds.guildName(GuildI) & FONTTYPE_INFO)
        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Status: Lider" & FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Oro: " & UserList(userindex).Stats.GLD & "  Posicion: " & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y & " en mapa " & UserList(userindex).Pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Dados: " & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) & FONTTYPE_INFO)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
With UserList(userindex)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & .name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & .clase & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & charName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||El pj no existe: " & charName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & UserList(userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).amount & FONTTYPE_INFO)
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
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & charName & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
            End If
        Next j
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & charName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| SkillLibres:" & UserList(userindex).Stats.SkillPts & FONTTYPE_INFO)
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
        If EsMascotaCiudadano Then Call SendData(SendTarget.ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)


'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            UserList(UserIndex).Reputacion.NobleRep = 0
            UserList(UserIndex).Reputacion.PlebeRep = 0
            UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200
            If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO
                If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.BandidoRep = MAXREP
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
End If

End Sub

Function PuedeApuñalar(ByVal userindex As Integer) As Boolean

If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(userindex).clase) = "ASESINO") And _
  (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer)

If UserList(userindex).flags.Hambre = 0 And _
   UserList(userindex).flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(userindex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(userindex).Stats.ELV > 3 _
        And UserList(userindex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(userindex).Stats.ELV >= 6 _
        And UserList(userindex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(userindex).Stats.ELV >= 10 _
        And UserList(userindex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = RandomNumber(1, Prob)
    
    Dim lvl As Integer
    lvl = UserList(userindex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(userindex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
        UserList(userindex).Stats.UserSkills(Skill) = UserList(userindex).Stats.UserSkills(Skill) + 1
        Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(userindex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
        
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + 50
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_FIGHT)
        Call CheckUserLevel(userindex)
    End If
End If

End Sub

Sub UserDie(ByVal userindex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(UserIndex).genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "QDL" & UserList(userindex).Char.CharIndex)
    
    UserList(userindex).Stats.MinHP = 0
    UserList(userindex).Stats.MinSta = 0
    UserList(userindex).flags.AtacadoPorNpc = 0
    UserList(userindex).flags.AtacadoPorUser = 0
    UserList(userindex).flags.Envenenado = 0
    UserList(userindex).flags.Muerto = 1
    
    
    Dim aN As Integer
    
    aN = UserList(userindex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<<< Paralisis >>>>
    If UserList(userindex).flags.Paralizado = 1 Then
        UserList(userindex).flags.Paralizado = 0
        Call SendData(SendTarget.ToIndex, userindex, 0, "PARADOK")
    End If
    
    '<<< Estupidez >>>
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, userindex, 0, "NESTUP")
    End If
    
    '<<<< Descansando >>>>
    If UserList(userindex).flags.Descansar Then
        UserList(userindex).flags.Descansar = False
        Call SendData(SendTarget.ToIndex, userindex, 0, "DOK")
    End If
    
    '<<<< Meditando >>>>
    If UserList(userindex).flags.Meditando Then
        UserList(userindex).flags.Meditando = False
        Call SendData(SendTarget.ToIndex, userindex, 0, "MEDOK")
    End If
    
    '<<<< Invisible >>>>
    If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
    End If
    
    If TriggerZonaPelea(userindex, userindex) <> TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(userindex) Or criminal(userindex) Then
            Call TirarTodo(userindex)
        Else
            If EsNewbie(userindex) Then Call TirarTodosLosItemsNoNewbies(userindex)
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
    End If
    'desequipar municiones
    If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(userindex).Char.loops = LoopAdEternum Then
        UserList(userindex).Char.FX = 0
        UserList(userindex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(userindex).flags.Mimetizado = 1 Then
        UserList(userindex).Char.body = UserList(userindex).CharMimetizado.body
        UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
        UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
        UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
        UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        UserList(userindex).Counters.Mimetismo = 0
        UserList(userindex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(userindex).flags.Navegando = 0 Then
        UserList(userindex).Char.body = iCuerpoMuerto
        UserList(userindex).Char.Head = iCabezaMuerto
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.CascoAnim = NingunCasco
    Else
        UserList(userindex).Char.body = iFragataFantasmal ';)
    End If
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(userindex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                    UserList(userindex).MascotasIndex(i) = 0
                    UserList(userindex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(userindex).NroMacotas = 0
    
    
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
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, val(userindex), UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call SendUserStatsBox(userindex)
    
    
    '<<Castigos por party>>
    If UserList(userindex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(userindex, UserList(userindex).Stats.ELV * -10 * mdParty.CantMiembros(userindex), UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
    End If

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
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

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(nPos.Map, tX, tY) Then
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

Sub WarpUserChar(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

    'Quitar el dialogo
    Call SendToUserArea(userindex, "QDL" & UserList(userindex).Char.CharIndex)
    Call SendData(SendTarget.ToIndex, userindex, UserList(userindex).Pos.Map, "QTDL")
    
    OldMap = UserList(userindex).Pos.Map
    OldX = UserList(userindex).Pos.X
    OldY = UserList(userindex).Pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, userindex)
    
    If OldMap <> Map Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "CM" & Map & "," & MapInfo(UserList(userindex).Pos.Map).MapVersion)
        Call SendData(SendTarget.ToIndex, userindex, 0, "TM" & MapInfo(Map).Music)
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(userindex).Pos.X = X
    UserList(userindex).Pos.Y = Y
    UserList(userindex).Pos.Map = Map
    
    Call MakeUserChar(SendTarget.ToMap, 0, Map, userindex, Map, X, Y)
    Call SendData(SendTarget.ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1) And (Not UserList(userindex).flags.AdminInvisible = 1) Then
        Call SendToUserArea(userindex, "NOVER" & UserList(userindex).Char.CharIndex & ",1", EncriptarProtocolosCriticos)
    End If
    
    If FX And UserList(userindex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).Pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    Call WarpMascotas(userindex)
End Sub

Sub UpdateUserMap(ByVal userindex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

'EnviarNoche UserIndex

On Error GoTo 0

Map = UserList(userindex).Pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).userindex > 0 And userindex <> MapData(Map, X, Y).userindex Then
            Call MakeUserChar(SendTarget.ToIndex, userindex, 0, MapData(Map, X, Y).userindex, Map, X, Y)
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                If UserList(MapData(Map, X, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).userindex).flags.Oculto = 1 Then Call SendCryptedData(SendTarget.ToIndex, userindex, 0, "NOVER" & UserList(MapData(Map, X, Y).userindex).Char.CharIndex & ",1")
            Else
#End If
                If UserList(MapData(Map, X, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).userindex).flags.Oculto = 1 Then Call SendData(SendTarget.ToIndex, userindex, 0, "NOVER" & UserList(MapData(Map, X, Y).userindex).Char.CharIndex & ",1")
#If SeguridadAlkon Then
            End If
#End If
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).ObjInfo, Map, X, Y)
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                End If
            End If
        End If
        
    Next X
Next Y

End Sub


Sub WarpMascotas(ByVal userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(userindex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Pierdes el control de tus mascotas." & FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(userindex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userindex).Pos, False, PetRespawn(i))
            UserList(userindex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(userindex).MascotasIndex(i) = 0 Then
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
                If UserList(userindex).NroMacotas > 0 Then UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
            Npclist(UserList(userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(userindex).MascotasIndex(i)).Target = 0
            Npclist(UserList(userindex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    UserList(userindex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal userindex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(userindex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(userindex).NroMacotas Then UserList(userindex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal userindex As Integer, Optional ByVal Tiempo As Integer = -1)
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
        UserList(userindex).Counters.Saliendo = True
        UserList(userindex).Counters.Salir = IIf(UserList(userindex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList(userindex).Pos.Map).Pk, 0, Tiempo)
        
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Cerrando...Se cerrará el juego en " & UserList(userindex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
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
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(userindex).flags.EstaEmpo = 0
    UserList(userindex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj Inexistente" & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Estadisticas de: " & Nombre & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD") & FONTTYPE_INFO)
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
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & charName & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & charName & FONTTYPE_INFO)
End If

End Sub

Sub VolverCriminal(ByVal userindex As Integer)
'**************************************************************
'Author: Unknow
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente
'**************************************************************
If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub

If UserList(userindex).flags.Privilegios < PlayerType.SemiDios Then
    UserList(userindex).Reputacion.BurguesRep = 0
    UserList(userindex).Reputacion.NobleRep = 0
    UserList(userindex).Reputacion.PlebeRep = 0
    UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + vlASALTO
    If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(userindex).Reputacion.BandidoRep = MAXREP
    If UserList(userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)
End If
Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex)
Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
End Sub

Sub VolverCiudadano(ByVal userindex As Integer)
'**************************************************************
'Author: Unknow
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 6 Then Exit Sub

UserList(userindex).Reputacion.LadronesRep = 0
UserList(userindex).Reputacion.BandidoRep = 0
UserList(userindex).Reputacion.AsesinoRep = 0
UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlASALTO
If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(userindex).Reputacion.PlebeRep = MAXREP
'Tenemos que actualizar el Tag del usuario. Esto no es lo optimo, ya que es un 1/0 que cambia en el paquete.
Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex)
Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, userindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
End Sub

