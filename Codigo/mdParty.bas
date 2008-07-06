Attribute VB_Name = "mdParty"
'**************************************************************
' mdParty.bas - Library of functions to manipulate parties.
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************


Option Explicit

''
' SOPORTES PARA LAS PARTIES
' (Ver este modulo como una clase abstracta "PartyManager")
'


''
'cantidad maxima de parties en el servidor
Public Const MAX_PARTIES As Integer = 300

''
'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 15

''
'Cantidad maxima de gente en la party
Public Const PARTY_MAXMEMBERS As Byte = 5

''
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = False

''
'maxima diferencia de niveles permitida en una party
Public Const MAXPARTYDELTALEVEL As Byte = 7

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 2

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 18

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

''
'Numero al que elevamos el nivel de cada miembro de la party
'Esto es usado para calcular la distribución de la experiencia entre los miembros
'Se lee del archivo de balance
Public ExponenteNivelParty As Single

''
'tPartyMember
'
' @param UserIndex UserIndex
' @param Experiencia Experiencia
'
Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Double
End Type


Public Function NextParty() As Integer
Dim i As Integer
NextParty = -1
For i = 1 To MAX_PARTIES
    If Parties(i) Is Nothing Then
        NextParty = i
        Exit Function
    End If
Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
    PuedeCrearParty = True
'    If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then
    
    If CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
        Call WriteConsoleMsg(UserIndex, "Tu carisma y liderazgo no son suficientes para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
Dim tInt As Integer
If UserList(UserIndex).PartyIndex = 0 Then
    If UserList(UserIndex).flags.Muerto = 0 Then
        If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear mas parties", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "La party está llena, no puedes entrar", FontTypeNames.FONTTYPE_PARTY)
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
                    UserList(UserIndex).PartyIndex = tInt
                    UserList(UserIndex).PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes hacerte líder.", FontTypeNames.FONTTYPE_PARTY)
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡ Te has convertido en líder de la party !", FontTypeNames.FONTTYPE_PARTY)
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(UserIndex, " No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
    End If
Else
    Call WriteConsoleMsg(UserIndex, " Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer

    If UserList(UserIndex).PartyIndex > 0 Then
        'si ya esta en una party
        Call WriteConsoleMsg(UserIndex, " Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, " ¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(UserIndex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
            Call WriteConsoleMsg(UserIndex, " El fundador decidirá si te acepta en la party", FontTypeNames.FONTTYPE_PARTY)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(tInt).name & " no es fundador de ninguna party.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(UserIndex, " Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY", FontTypeNames.FONTTYPE_PARTY)
        UserList(UserIndex).PartySolicitud = 0
    End If
End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    If Parties(PI).SaleMiembro(UserIndex) Then
        'sale el leader
        Set Parties(PI) = Nothing
    Else
        UserList(UserIndex).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(UserIndex, " No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer
PI = UserList(leader).PartyIndex
If PI > 0 Then
    If PI = UserList(OldMember).PartyIndex Then
        If Parties(PI).EsPartyLeader(leader) Then
            If Parties(PI).SaleMiembro(OldMember) Then
                'si la funcion me da true, entonces la party se disolvio
                'y los partyindex fueron reseteados a 0
                Set Parties(PI) = Nothing
            Else
                UserList(OldMember).PartyIndex = 0
            End If
        Else
            Call WriteConsoleMsg(leader, " Solo el fundador puede expulsar miembros de una party.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(leader, UserList(OldMember).name & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(leader, " No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub


Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(leader).PartyIndex

If PI > 0 Then
    If Parties(PI).EsPartyLeader(leader) Then
        If UserList(NewMember).PartyIndex = 0 Then
            If Not UserList(leader).flags.Muerto = 1 Then
                If Not UserList(NewMember).flags.Muerto = 1 Then
                    If UserList(NewMember).PartySolicitud = PI Then
                        If Parties(PI).PuedeEntrar(NewMember, razon) Then
                            If Parties(PI).NuevoMiembro(NewMember) Then
                                Call Parties(PI).MandarMensajeAConsola(UserList(leader).name & " ha aceptado a " & UserList(NewMember).name & " en la party.", "Servidor")
                                UserList(NewMember).PartyIndex = PI
                                UserList(NewMember).PartySolicitud = 0
                            Else
                                'no pudo entrar
                                'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                                Call SendData(SendTarget.ToAdmins, leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_PARTY))
                            End If
                        Else
                            'no debe entrar
                            Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_PARTY)
                        End If
                    Else
                        Call WriteConsoleMsg(leader, UserList(NewMember).name & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                        Exit Sub
                    End If
                Else
                    Call WriteConsoleMsg(leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(leader, "¡Estás muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(leader, UserList(NewMember).name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_PARTY)
            ' ya tiene party el otro tipo
        End If
    Else
        Call WriteConsoleMsg(leader, "No eres líder, no puedes aceptar miembros.", FontTypeNames.FONTTYPE_PARTY)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(leader, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).name)
    End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
Dim PI As Integer
Dim texto As String

    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        Call WriteConsoleMsg(UserIndex, texto, FontTypeNames.FONTTYPE_PARTY)
    End If
    

End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI > 0 Then
    If PI = UserList(NewLeader).PartyIndex Then
        If UserList(NewLeader).flags.Muerto = 0 Then
            If Parties(PI).EsPartyLeader(OldLeader) Then
                If Parties(PI).HacerLeader(NewLeader) Then
                    Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).name, UserList(OldLeader).name)
                Else
                    Call WriteConsoleMsg(OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
                End If
            Else
                Call WriteConsoleMsg(OldLeader, "¡No eres el líder!", FontTypeNames.FONTTYPE_PARTY)
            End If
        Else
            Call WriteConsoleMsg(OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, UserList(NewLeader).name & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub


Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim i As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To MAX_PARTIES
        If Not Parties(i) Is Nothing Then
            Call Parties(i).FlushExperiencia
        End If
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False

End If

End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, X As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)


End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

End Function
