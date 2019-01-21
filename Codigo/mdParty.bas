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
'Esto es usado para calcular la distribuci�n de la experiencia entre los miembros
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
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

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
'***************************************************
'Author: Unknown
'Last Modification: 05/22/2010 (Marco)
' - 05/22/2010 : staff members aren't allowed to party anyone. (Marco)
'***************************************************
    
    PuedeCrearParty = True
    
    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) = 0 Then
    'staff members aren't allowed to party anyone.
        Call WriteConsoleMsg(UserIndex, "�Los miembros del staff no pueden crear partys!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
        Call WriteConsoleMsg(UserIndex, "Tu carisma y liderazgo no son suficientes para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim tInt As Integer

With UserList(UserIndex)
    If .PartyIndex = 0 Then
        If .flags.Muerto = 0 Then
            If .Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
                tInt = mdParty.NextParty
                If tInt = -1 Then
                    Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear m�s parties.", FontTypeNames.FONTTYPE_PARTY)
                    Exit Sub
                Else
                    Set Parties(tInt) = New clsParty
                    If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "La party est� llena, no puedes entrar.", FontTypeNames.FONTTYPE_PARTY)
                        Set Parties(tInt) = Nothing
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "�Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
                        .PartyIndex = tInt
                        .PartySolicitud = 0
                        If Not Parties(tInt).HacerLeader(UserIndex) Then
                            Call WriteConsoleMsg(UserIndex, "No puedes hacerte l�der.", FontTypeNames.FONTTYPE_PARTY)
                        Else
                            Call WriteConsoleMsg(UserIndex, "�Te has convertido en l�der de la party!", FontTypeNames.FONTTYPE_PARTY)
                        End If
                    End If
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
            End If
        Else
            'Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_PARTY)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
    End If
End With
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/22/2010 (Marco)
' - 05/22/2010 : staff members aren't allowed to party anyone. (Marco)
'18/09/2010: ZaMa - Ahora le avisa al funda de la party cuando alguien quiere ingresar a la misma.
'18/09/2010: ZaMa - Contemple mas ecepciones (solo se le puede mandar party al lider)
'***************************************************

    'ESTO ES enviado por el PJ para solicitar el ingreso a la party
    Dim TargetUserIndex As Integer
    Dim PartyIndex As Integer

    With UserList(UserIndex)
    
        'staff members aren't allowed to party anyone
        If (.flags.Privilegios And PlayerType.User) = 0 Then
            Call WriteConsoleMsg(UserIndex, "�Los miembros del staff no pueden unirse a partys!", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        If .PartyIndex > 0 Then
            'si ya esta en una party
            Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
            .PartySolicitud = 0
            Exit Sub
        End If
        
        ' Muerto?
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            .PartySolicitud = 0
            Exit Sub
        End If
        
        TargetUserIndex = .flags.TargetUser
        ' Target valido?
        If TargetUserIndex > 0 Then
        
            PartyIndex = UserList(TargetUserIndex).PartyIndex
            ' Tiene party?
            If PartyIndex > 0 Then
            
                ' Es el lider?
                If Parties(PartyIndex).EsPartyLeader(TargetUserIndex) Then
                    .PartySolicitud = PartyIndex
                    Call WriteConsoleMsg(UserIndex, "El lider decidir� si te acepta en la party.", FontTypeNames.FONTTYPE_PARTY)
                    Call WriteConsoleMsg(TargetUserIndex, .Name & " solicita ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                
                ' No es lider
                Else
                    Call WriteConsoleMsg(UserIndex, UserList(TargetUserIndex).Name & " no es lider de la party.", FontTypeNames.FONTTYPE_PARTY)
                End If
            
            ' No tiene party
            Else
                Call WriteConsoleMsg(UserIndex, UserList(TargetUserIndex).Name & " no pertenece a ninguna party.", FontTypeNames.FONTTYPE_PARTY)
                .PartySolicitud = 0
                Exit Sub
            End If
        
        ' Target inv�lido
        Else
            Call WriteConsoleMsg(UserIndex, "Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY", FontTypeNames.FONTTYPE_PARTY)
            .PartySolicitud = 0
        End If
        
    End With

End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

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
    Call WriteConsoleMsg(UserIndex, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim PI As Integer
PI = UserList(leader).PartyIndex

If PI = UserList(OldMember).PartyIndex Then
    If Parties(PI).SaleMiembro(OldMember) Then
        'si la funcion me da true, entonces la party se disolvio
        'y los partyindex fueron reseteados a 0
        Set Parties(PI) = Nothing
    Else
        UserList(OldMember).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(OldMember).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

''
' Determines if a user can use party commands like /acceptparty or not.
'
' @param User Specifies reference to user
' @return  True if the user can use party commands, false if not.
Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
'*************************************************
'Author: Marco Vanotti(Marco)
'Last modified: 05/05/09
'
'*************************************************
    Dim PI As Integer
    
    PI = UserList(User).PartyIndex
    
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "�No eres el l�der de tu party!", FontTypeNames.FONTTYPE_PARTY)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
End Function

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa al lider si intenta aceptar a alguien que sea mimebro de su propia party.
'***************************************************

    'el UI es el leader
    Dim PI As Integer
    Dim razon As String
    
    PI = UserList(leader).PartyIndex
    
    With UserList(NewMember)
        If .PartySolicitud = PI Then
            If Not .flags.Muerto = 1 Then
                If .PartyIndex = 0 Then
                    If Parties(PI).PuedeEntrar(NewMember, razon) Then
                        If Parties(PI).NuevoMiembro(NewMember) Then
                            Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & _
                                " ha aceptado a " & .Name & " en la party.", "Servidor")
                            .PartyIndex = PI
                            .PartySolicitud = 0
                        Else
                            'no pudo entrar
                            'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                            Call SendData(SendTarget.ToAdmins, leader, _
                                PrepareMessageConsoleMsg(" Servidor> CAT�STROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", _
                                FontTypeNames.FONTTYPE_PARTY))
                            End If
                        Else
                        'no debe entrar
                        Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_PARTY)
                    End If
                Else
                    If .PartyIndex = PI Then
                        Call WriteConsoleMsg(leader, LCase(.Name) & _
                            " ya es miembro de la party.", FontTypeNames.FONTTYPE_PARTY)
                    Else
                        Call WriteConsoleMsg(leader, .Name & " ya es miembro de otra party.", _
                            FontTypeNames.FONTTYPE_PARTY)
                    End If
                    
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(leader, "�Est� muerto, no puedes aceptar miembros en ese estado!", _
                    FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            End If
        Else
            If .PartyIndex = PI Then
                Call WriteConsoleMsg(leader, LCase(.Name) & _
                    " ya es miembro de la party.", FontTypeNames.FONTTYPE_PARTY)
            Else
                Call WriteConsoleMsg(leader, LCase(.Name) & _
                    " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
            End If
            
            Exit Sub
        End If
    End With
    
End Sub

Private Function IsPartyMember(ByVal UserIndex As Integer, ByVal PartyIndex As Integer)
    Dim MemberIndex As Integer
    
    For MemberIndex = 1 To PARTY_MAXMEMBERS
        
    Next MemberIndex
End Function

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
    End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/27/09 (Budi)
'Adapte la funci�n a los nuevos m�todos de clsParty
'*************************************************
Dim i As Integer
Dim PI As Integer
Dim Text As String
Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer

    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
        Text = "Nombre(Exp): "
        For i = 1 To PARTY_MAXMEMBERS
            If MembersOnline(i) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
            End If
        Next i
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_PARTY)
    End If
    
End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI = UserList(NewLeader).PartyIndex Then
    If UserList(NewLeader).flags.Muerto = 0 Then
        If Parties(PI).HacerLeader(NewLeader) Then
            Call Parties(PI).MandarMensajeAConsola("El nuevo l�der de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
        Else
            Call WriteConsoleMsg(OldLeader, "�No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub


Public Sub ActualizaExperiencias()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

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
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)


End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

End Function

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param UserInidex Specifies reference to user
' @remarks When a user level up and he is in a party, we call this sub to don't desestabilice the party exp formula
Public Sub ActualizarSumaNivelesElevados(ByVal UserIndex As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 28/10/08
'
'*************************************************
    If UserList(UserIndex).PartyIndex > 0 Then
        Call Parties(UserList(UserIndex).PartyIndex).UpdateSumaNivelesElevados(UserList(UserIndex).Stats.ELV)
    End If
End Sub


