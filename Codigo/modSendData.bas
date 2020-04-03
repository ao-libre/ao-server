Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martin Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs

End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus) - Rewrite of original
    'Last Modify Date: 14/11/2010
    'Last modified by: ZaMa
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '**************************************************************
    On Error Resume Next

    Dim LoopC As Long
    
    Select Case sndRoute

        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

                        ' Denounces can be desactivated
                        If IsDenounce Then
                            
                            If UserList(LoopC).flags.SendDenounces Then
                                Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                            End If

                        Else
                            Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                        End If

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToAll

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToAllButIndex

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToConsejo

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToConsejoCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRolesMasters

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If Not criminal(LoopC) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminales

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If criminal(LoopC) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToReal

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRealYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToHigherAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
            
        Case SendTarget.ToGMsAreaButRmsOrCounselors
            Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub

        Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
            Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
            Exit Sub

    End Select

End Sub

Private Sub SendToUserArea(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToUserAreaButindex(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If tempIndex <> Userindex Then
            If EstanMismoArea(Userindex, tempIndex) Then
                If UserList(tempIndex).ConnIDValida Then
                    Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
                End If
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToDeadUserArea(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            'Dead and admins read
            If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToUserGuildArea(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    If UserList(Userindex).GuildIndex = 0 Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)
        
        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList(Userindex).GuildIndex Or ((UserList(tempIndex).flags.Privilegios And PlayerType.Dios) And (UserList(tempIndex).flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToUserPartyArea(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub
    
    If UserList(Userindex).PartyIndex = 0 Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida And UserList(tempIndex).PartyIndex = UserList(Userindex).PartyIndex Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal Userindex As Integer, _
                                          ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida Then
                If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                    Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
                End If
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = Npclist(NpcIndex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoAreaNPC(NpcIndex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoAreaPos(tempIndex, X, Y) Then
            If UserList(tempIndex).ConnIDValida Then
                Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
            End If
        End If

    Next LoopC

End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim LoopC     As Long

    Dim tempIndex As Integer
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next LoopC

End Sub

Public Sub SendToMapButIndex(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim LoopC     As Long

    Dim Map       As Integer

    Dim tempIndex As Integer
    
    Map = UserList(Userindex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)
        
        If tempIndex <> Userindex And UserList(tempIndex).ConnIDValida Then
            Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next LoopC

End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal Userindex As Integer, _
                                            ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 12/02/2010
    '12/02/2010: ZaMa - Restrinjo solo a dioses, admins y gms.
    '15/02/2010: ZaMa - Cambio el nombre de la funcion (viejo: ToGmsArea, nuevo: ToGmsAreaButRMsOrCounselors)
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)
        
        With UserList(tempIndex)

            If EstanMismoArea(Userindex, tempIndex) Then
                If .ConnIDValida Then
                    ' Exclusivo para dioses, admins y gms
                    If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not PlayerType.RoleMaster) = .flags.Privilegios Then
                        Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
                    End If
                End If
            End If

        End With

    Next LoopC

End Sub

Private Sub SendToUsersAreaButGMs(ByVal Userindex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida Then
                If UserList(tempIndex).flags.Privilegios And PlayerType.User Then
                    Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
                End If
            End If
        End If

    Next LoopC

End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal Userindex As Integer, _
                                                     ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer

    Map = UserList(Userindex).Pos.Map

    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)

        If EstanMismoArea(Userindex, tempIndex) Then
            If UserList(tempIndex).ConnIDValida Then
                If UserList(tempIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                    Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(sdData)
                End If
            End If
        End If

    Next LoopC

End Sub

Public Sub AlertarFaccionarios(ByVal Userindex As Integer)

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Alerta a los faccionarios, dandoles una orientacion
    '**************************************************************
    Dim LoopC     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer
    Dim Font      As FontTypeNames
    Dim tempData  As String
    
    If esCaos(Userindex) Then
        Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
    Else
        Font = FontTypeNames.FONTTYPE_CONSEJO
    End If
    
    Map = UserList(Userindex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).Count()
        tempIndex = ConnGroups(Map).Item(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            
            If tempIndex <> Userindex Then

                ' Solo se envia a los de la misma faccion
                If SameFaccion(Userindex, tempIndex) Then
                
                    tempData = PrepareMessageConsoleMsg("Escuchas el llamado de un companero que proviene del " & GetDireccion(Userindex, tempIndex), Font)
                    
                    Call UserList(tempIndex).outgoingData.WriteASCIIStringFixed(tempData)

                End If

            End If

        End If

    Next LoopC

End Sub

