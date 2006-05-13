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
    If (Not Criminal(VictimIndex)) Then
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


Sub RevivirUsuario(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 35

'No puede estar empollando
UserList(UserIndex).flags.EstaEmpo = 0
UserList(UserIndex).EmpoCont = 0

If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Heading = Heading
    UserList(UserIndex).Char.WeaponAnim = Arma
    UserList(UserIndex).Char.ShieldAnim = Escudo
    UserList(UserIndex).Char.CascoAnim = Casco
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(UserIndex, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)
    End If
End Sub

Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
       cad = cad & UserList(UserIndex).Stats.UserSkills(i) & ","
    Next i
    
    SendData SendTarget.ToIndex, UserIndex, 0, "SKILLS" & cad$
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim cad As String
    
    cad = cad & UserList(UserIndex).Reputacion.AsesinoRep & ","
    cad = cad & UserList(UserIndex).Reputacion.BandidoRep & ","
    cad = cad & UserList(UserIndex).Reputacion.BurguesRep & ","
    cad = cad & UserList(UserIndex).Reputacion.LadronesRep & ","
    cad = cad & UserList(UserIndex).Reputacion.NobleRep & ","
    cad = cad & UserList(UserIndex).Reputacion.PlebeRep & ","
    
    Dim L As Long
    
    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
        (-UserList(UserIndex).Reputacion.BandidoRep) + _
        UserList(UserIndex).Reputacion.BurguesRep + _
        (-UserList(UserIndex).Reputacion.LadronesRep) + _
        UserList(UserIndex).Reputacion.NobleRep + _
        UserList(UserIndex).Reputacion.PlebeRep
    L = L / 6
    
    UserList(UserIndex).Reputacion.Promedio = L
    
    cad = cad & UserList(UserIndex).Reputacion.Promedio
    
    SendData SendTarget.ToIndex, UserIndex, 0, "FAMA" & cad
End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad As String
For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call SendData(SendTarget.ToIndex, UserIndex, 0, "ATR" & cad)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEST" & .Faccion.CiudadanosMatados & "," & _
                .Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
                .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena)
End With

End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(UserIndex, "BP" & UserList(UserIndex).Char.CharIndex)
        Call QuitarUser(UserIndex, UserList(UserIndex).Pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BP" & UserList(UserIndex).Char.CharIndex)
    End If
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If
        
        'Place character on map
        MapData(Map, X, Y).UserIndex = UserIndex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(UserIndex).GuildIndex > 0 Then
            klan = Guilds(UserList(UserIndex).GuildIndex).GuildName
        End If
        
        Dim bCr As Byte
        Dim SendPrivilegios As Byte
       
        bCr = Criminal(UserIndex)

        If klan <> "" Then
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        Else
                            'Hide the name and clan
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 4, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        Else
                            'Hide the name and clan
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        End If
                    Else
                        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 4, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map)
                Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
            End If
        Else 'if tiene clan
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & "," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        Else
                            'Hide the name
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 4, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & "," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        Else
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & ",," & bCr & "," & IIf(UserList(UserIndex).flags.EsRolesMaster, 5, UserList(UserIndex).flags.Privilegios))
                        End If
                    Else
                        Call SendData(SendTarget.ToIndex, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 4, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map)
                Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
            End If
       End If   'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU
    
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_NIVEL)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    If UserList(UserIndex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 5
    End If
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
       
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
    
    If Not EsNewbie(UserIndex) And WasNewbie Then
        Call QuitarNewbieObj(UserIndex)
        If UCase$(MapInfo(UserList(UserIndex).Pos.Map).Restringir) = "SI" Then
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes abandonar el Dungeon Newbie." & FONTTYPE_WARNING)
        End If
    End If

    If UserList(UserIndex).Stats.ELV < 11 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    ElseIf UserList(UserIndex).Stats.ELV < 25 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    Else
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.2
    End If

    Dim AumentoHP As Integer
    Select Case UCase$(UserList(UserIndex).Clase)
        Case "GUERRERO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 12)
                Case 19, 18
                    AumentoHP = RandomNumber(8, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19, 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select

            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PIRATA"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 18, 19
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19, 18
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "LADRON"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case "MAGO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case "LEÑADOR"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case "MINERO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case "PESCADOR"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case "CLERIGO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 11)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "DRUIDA"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case Else
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select

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
    If AumentoHP > 0 Then SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoSTA > 0 Then SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoSTA & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData SendTarget.ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData SendTarget.ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    Call LogDesarrollo(Date & " " & UserList(UserIndex).name & " paso a nivel " & UserList(UserIndex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, Pts)
   
    SendUserStatsBox UserIndex
    
Loop
'End If


Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
    
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
        If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...
#If SeguridadAlkon Then
            Call SendCryptedMoveChar(nPos.Map, UserIndex, nPos.X, nPos.Y)
#Else
            Call SendToUserAreaButindex(UserIndex, "+" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
#End If
        End If
        
        'Update map and user pos
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(UserIndex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ 3)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")
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

Sub SendUserStatsBox(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp)
End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    GuildI = UserList(UserIndex).GuildIndex
    If GuildI > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clan: " & Guilds(GuildI).GuildName & FONTTYPE_INFO)
        If UCase$(Guilds(GuildI).GetLeader) = UCase$(UserList(sendIndex).name) Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Status: Lider" & FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) & FONTTYPE_INFO)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & .name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & .Clase & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(UserIndex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
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
        Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & UserList(UserIndex).name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| SkillLibres:" & UserList(UserIndex).Stats.SkillPts & FONTTYPE_INFO)
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
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
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

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(UserIndex).Clase) = "ASESINO") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

If UserList(UserIndex).flags.Hambre = 0 And _
   UserList(UserIndex).flags.Sed = 0 Then
    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(UserIndex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(UserIndex).Stats.ELV > 3 _
        And UserList(UserIndex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(UserIndex).Stats.ELV >= 6 _
        And UserList(UserIndex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(UserIndex).Stats.ELV >= 10 _
        And UserList(UserIndex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = RandomNumber(1, Prob)
    
    Dim lvl As Integer
    lvl = UserList(UserIndex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
        UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
        
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + 50
        If UserList(UserIndex).Stats.Exp > MAXEXP Then _
            UserList(UserIndex).Stats.Exp = MAXEXP
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_FIGHT)
        Call CheckUserLevel(UserIndex)
    End If
End If

End Sub

Sub UserDie(ByVal UserIndex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(UserIndex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1
    
    
    Dim aN As Integer
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOK")
    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")
    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
    End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
    End If
    
    If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
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
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
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
    If UserList(UserIndex).Char.loops = LoopAdEternum Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
        UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal ';)
    End If
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                    UserList(UserIndex).MascotasIndex(i) = 0
                    UserList(UserIndex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(UserIndex).NroMacotas = 0
    
    
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
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call SendUserStatsBox(UserIndex)
    
    
    '<<Castigos por party>>
    If UserList(UserIndex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(UserIndex, UserList(UserIndex).Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    End If

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.RecompensasReal = 0
        End If
        
        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
            'con esto evitamos que se vuelva a reenlistar
        End If
    Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then _
                UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = 0
            UserList(Atacante).Faccion.RecompensasCaos = 0
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
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
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

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

    'Quitar el dialogo
    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.CharIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")
    
    OldMap = UserList(UserIndex).Pos.Map
    OldX = UserList(UserIndex).Pos.X
    OldY = UserList(UserIndex).Pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, UserIndex)
        
    If OldMap <> Map Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.Map = Map
    
    Call MakeUserChar(SendTarget.ToMap, 0, Map, UserIndex, Map, X, Y)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendToUserArea(UserIndex, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1", EncriptarProtocolosCriticos)
    End If
    
    If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    Call WarpMascotas(UserIndex)
End Sub

Sub UpdateUserMap(ByVal UserIndex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

'EnviarNoche UserIndex

On Error GoTo 0

Map = UserList(UserIndex).Pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
            Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call SendCryptedData(SendTarget.ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
            Else
#End If
                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
#If SeguridadAlkon Then
            End If
#End If
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                End If
            End If
        End If
        
    Next X
Next Y

End Sub


Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(UserIndex).NroMacotas
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
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Pierdes el control de tus mascotas." & FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
            UserList(UserIndex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(UserIndex).MascotasIndex(i) = 0 Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
        End If
    Next i
    
    UserList(UserIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal Tiempo As Integer = -1)
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList(UserIndex).Pos.Map).Pk, 0, Tiempo)
        
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
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
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
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
Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(SendTarget.ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub
