Attribute VB_Name = "Mod_ClanvsClan"
Option Explicit

'maTih.-

Public Type cvc_User

    en_CVC         As Boolean
    cvc_Target     As Integer
    cvc_MaxUsers   As Byte

End Type

Type cvc_Clanes

    Guild_Index    As Integer
    Num_Users      As Byte
    UsUaRiOs()     As Integer
    Rounds         As Byte

End Type

Type cvc_Data

    Guild(1 To 2)  As cvc_Clanes
    cvc_Enabled    As Boolean
    count_Down     As Byte
    max_Users      As Byte

End Type

Public CVC_Info     As cvc_Data

Public usersClan1   As Byte
Public usersClan2   As Byte

Public menorCant    As Byte

'constantes de pos..
Const PRIMER_CLAN_X As Byte = 39
Const SECOND_CLAN_X As Byte = 20
Const PRIMER_CLAN_Y As Byte = 78
Const SECOND_CLAN_Y As Byte = 74
Const MAPA_CVC      As Integer = 273

Public Sub Enviar(ByVal Userindex As Integer, _
                  ByVal targetIndex As Integer) ', ByVal max_Users_Guild As Byte)

    '
    ' @ Envia una solicitud.

    With UserList(Userindex)

        Dim other_Guild As Integer

        Dim my_Guild    As Integer

        'Dim cant_Users  As Integer
                
        other_Guild = UserList(targetIndex).GuildIndex
        my_Guild = .GuildIndex
        usersClan1 = guilds(.GuildIndex).CantidadDeMiembros
        usersClan2 = guilds(other_Guild).CantidadDeMiembros

        If usersClan1 <= usersClan2 Then
            menorCant = usersClan1
        Else
            menorCant = usersClan2

        End If

        'seteo el target al otro usuario
        .cvcUser.cvc_Target = targetIndex
        .cvcUser.cvc_MaxUsers = menorCant
        UserList(targetIndex).cvcUser.cvc_Target = Userindex
                
        Call Protocol.WriteConsoleMsg(targetIndex, "El clan " & modGuilds.GuildName(my_Guild) & " desafia tu clan a un duelo de modalidad Clan vs Clan, si aceptas hazle click y tipea /ACVC.", FontTypeNames.FONTTYPE_GUILD)
        Call Protocol.WriteConsoleMsg(targetIndex, "La cantidad maxima de usuarios por clan es de : " & CStr(menorCant) & ".", FontTypeNames.FONTTYPE_GUILD)
        Call Protocol.WriteConsoleMsg(Userindex, "Ahora debes esperar que el lider acepte.", FontTypeNames.FONTTYPE_GUILD)

    End With

End Sub
 
Public Sub Aceptar(ByVal Userindex As Integer, ByVal targetIndex As Integer)

    '
    ' @ Acepta cvc.

    With UserList(Userindex)

        'clickio a quien le envio?

        If (targetIndex = .cvcUser.cvc_Target) Then
            Call Iniciar(targetIndex, Userindex, .GuildIndex, UserList(targetIndex).GuildIndex, UserList(targetIndex).cvcUser.cvc_MaxUsers)
        Else
            Call Protocol.WriteConsoleMsg(Userindex, UserList(targetIndex).Name & " no solicito ningun Clan vs Clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub
 
Private Sub Iniciar(ByVal userSend As Integer, _
                    ByVal userAccept As Integer, _
                    ByVal Guild_Desafiado As Integer, _
                    ByVal Guild_Desafiante As Integer, _
                    ByVal max_Users As Byte)

    '
    ' @ Inicia un nuevo cvc.

    With CVC_Info
        .cvc_Enabled = True
        'lleno los indices
        .Guild(1).Guild_Index = Guild_Desafiante
        .Guild(2).Guild_Index = Guild_Desafiado
        'redimensiono el array de usuarios
        ReDim .Guild(1).UsUaRiOs(1 To max_Users) As Integer
        ReDim .Guild(2).UsUaRiOs(1 To max_Users) As Integer
        .max_Users = max_Users

        'lleno el array de usuarios invalidos

        Dim j As Long

        For j = 1 To max_Users
            .Guild(1).UsUaRiOs(j) = -1
            .Guild(2).UsUaRiOs(j) = -1

        Next j

        For j = 1 To 2
            Call SendData(SendTarget.ToGuildMembers, .Guild(j).Guild_Index, Protocol.PrepareMessageConsoleMsg("CLAN VS CLAN > " & modGuilds.GuildName(.Guild(1).Guild_Index) & " vs " & modGuilds.GuildName(.Guild(2).Guild_Index) & " cada clan con " & CStr(.max_Users) & " Participantes, para participar tipea /IRCVC.", FontTypeNames.FONTTYPE_GUILD))

        Next j

        'lleno los primeros ui.
        Call ConectarCVC(userSend, 1)
        Call ConectarCVC(userAccept, 2)

    End With

End Sub
 
Private Sub EnviarMensajeCVC(ByVal to_Guild As Byte, ByRef send_Msg As String)

    '
    ' @ Envia un mensaje a un clan en cvc o a todos en el mismo.

    With CVC_Info

        Dim j As Long

        Dim i As Long

        'a todo el cvc.

        If (to_Guild = 0) Then

            For j = 1 To 2

                With .Guild(j)

                    For i = 1 To UBound(.UsUaRiOs())

                        If .UsUaRiOs(i) <> -1 Then
                            If UserList(.UsUaRiOs(i)).ConnID <> -1 Then
                                Call Protocol.WriteConsoleMsg(.UsUaRiOs(i), send_Msg, FontTypeNames.FONTTYPE_GUILD)

                            End If

                        End If

                    Next i

                End With

            Next j

            Exit Sub

        End If

        'a un solo clan

        For j = 1 To UBound(.Guild(to_Guild).UsUaRiOs())

            With .Guild(to_Guild)

                For i = 1 To UBound(.UsUaRiOs())

                    If .UsUaRiOs(i) <> -1 Then
                        If UserList(.UsUaRiOs(i)).ConnID <> -1 Then
                            Call Protocol.WriteConsoleMsg(.UsUaRiOs(i), send_Msg, FontTypeNames.FONTTYPE_GUILD)

                        End If

                    End If

                Next i

            End With

        Next j

    End With

End Sub
 
Public Sub MuereCVC(ByVal Userindex As Integer)

    '
    ' @ Muere un usuario en cvc.

    Dim num_Muertos As Byte

    Dim guild_Num   As Byte

    Dim guild_Win   As Byte

    guild_Num = Find_Guild_Num(UserList(Userindex).GuildIndex)
    num_Muertos = Get_Num_Dies(guild_Num)

    'murieron todos?

    If (num_Muertos >= CVC_Info.max_Users) Then
        If (guild_Num) = 1 Then
            guild_Win = 2
        Else
            guild_Win = 1

        End If

        'sumo el contador de rounds
        CVC_Info.Guild(guild_Win).Rounds = CVC_Info.Guild(guild_Win).Rounds + 1

        'cuantos rounds gano el equipo ganador?
        If (CVC_Info.Guild(guild_Win).Rounds >= 1) Then
            'gana el cvc
            Call GanaCVC(guild_Win, guild_Num)
        Else 'van empatados o gana 1 a 0.
            Call ReiniciarCVC

        End If

    End If

End Sub
 
Private Sub ReiniciarCVC()

    '
    ' @ Comienza otro duelo.

    Dim j  As Long

    Dim i  As Long

    Dim sX As Byte

    Dim sY As Byte

    With CVC_Info

        For i = 1 To 2

            With .Guild(i)

                For j = 1 To UBound(.UsUaRiOs())

                    If (.UsUaRiOs(j) <> -1) Then
                        If (UserList(.UsUaRiOs(j)).ConnID <> -1) Then
                            Call Get_Pos_By_Guild(.UsUaRiOs(j), CByte(i), sX, sY)

                            If (sX <> 0) And (sY <> 0) Then
                                'warp
                                Call WarpUserChar(.UsUaRiOs(j), MAPA_CVC, sX, sY, True)

                            End If

                        End If

                    End If

                Next j

            End With

        Next i

        Dim ganando_name As String

        Dim prepare_Text As String

        prepare_Text = "Victoria parcial para : "

        'el clan 1 tiene mas rounds ganadoss

        If (.Guild(1).Rounds > .Guild(2).Rounds) Then
            ganando_name = modGuilds.GuildName(.Guild(1).Guild_Index)
            prepare_Text = prepare_Text & ganando_name
        ElseIf (.Guild(2).Rounds > .Guild(1).Rounds) Then 'gana el equipo 2
            ganando_name = modGuilds.GuildName(.Guild(2).Guild_Index)
            prepare_Text = prepare_Text & ganando_name
        ElseIf (.Guild(1).Rounds = .Guild(2).Rounds) Then 'empate..
            prepare_Text = "Empate."

        End If

        Call EnviarMensajeCVC(0, "Resultado parcial : " & .Guild(1).Rounds & " - " & .Guild(2).Rounds)
        Call EnviarMensajeCVC(0, prepare_Text)

    End With

End Sub
 
Private Sub EraseCVC()

    '
    ' @ Limpia los datos del cvc

    With CVC_Info

        Dim j As Long

        Dim i As Long

        For j = 1 To 2
            For i = 1 To UBound(.Guild(j).UsUaRiOs())
                .Guild(j).UsUaRiOs(i) = -1

            Next i

            .Guild(j).Num_Users = 0
            .Guild(j).Guild_Index = 0
            .Guild(j).Rounds = 0

        Next j

        .cvc_Enabled = True
        .count_Down = 0
        .max_Users = 0
                
        Erase .Guild(1).UsUaRiOs
        Erase .Guild(2).UsUaRiOs

    End With

End Sub
 
Private Sub GanaCVC(ByVal guildWinner As Byte, ByVal guildLooser As Byte)

    '
    ' @ Termina el cvc.

    With CVC_Info

        Dim j      As Long

        Dim i      As Long

        Dim startX As Byte

        Dim startY As Byte

        startX = Ullathorpe.x
        startY = Ullathorpe.Y

        Dim sMessage As String

        sMessage = "CVC > Terminado."

        For i = 1 To 2
            For j = 1 To UBound(.Guild(i).UsUaRiOs())

                With .Guild(i)

                    'hay un user

                    If (.UsUaRiOs(j) <> -1) Then
                        'busco una pos
                        Call FindLegalPos(.UsUaRiOs(j), Ullathorpe.Map, CInt(startX), CInt(startY))

                        'hay pos

                        If (startX <> 0) And (startY <> 0) Then
                            Call WarpUserChar(.UsUaRiOs(j), Ullathorpe.Map, startX, startY, True)

                        End If

                        'reseteo los flags del user
                        UserList(.UsUaRiOs(j)).cvcUser.cvc_MaxUsers = 0
                        UserList(.UsUaRiOs(j)).cvcUser.cvc_Target = 0
                        UserList(.UsUaRiOs(j)).cvcUser.en_CVC = False

                    End If

                End With

            Next j
        Next i

        sMessage = sMessage & vbNewLine
        sMessage = modGuilds.GuildName(.Guild(guildWinner).Guild_Index) & " vencio a " & modGuilds.GuildName(.Guild(guildLooser).Guild_Index) & " en un duelo " & CStr(.max_Users) & " vs " & CStr(.max_Users) & "."
        Call SendData(SendTarget.ToAll, 0, Protocol.PrepareMessageConsoleMsg(sMessage, FontTypeNames.FONTTYPE_GUILD))
        'limpio la data
        Call EraseCVC

    End With

End Sub
 
Public Sub ConectarCVC(ByVal Userindex As Integer, _
                       Optional ByVal check_Enabled As Boolean = False)

    '
    ' @ Conecta un usuario al CVC.

    Dim guild_Num As Byte

    'hay qe aser los checkeos?

    If (check_Enabled = True) Then

        Dim ref_Error As String

        'no puede entrar ,lo informo.

        If (Can_Ingress(Userindex, ref_Error) = False) Then
            Call Protocol.WriteConsoleMsg(Userindex, ref_Error, FontTypeNames.FONTTYPE_GUILD)

            Exit Sub

        End If

        guild_Num = Find_Guild_Num(UserList(Userindex).GuildIndex)

    End If

    With CVC_Info.Guild(guild_Num)

        Dim user_Gindex As Byte

        Dim user_toPosX As Byte

        Dim user_toPosY As Byte

        user_Gindex = index_In_CVC(guild_Num)

        'tenemos un slot

        If (user_Gindex <> 0) Then
            .UsUaRiOs(user_Gindex) = Userindex
            Call EnviarMensajeCVC(0, UserList(Userindex).Name & " ingreso al cvc para el clan " & modGuilds.GuildName(UserList(Userindex).GuildIndex) & "!")
            'buscamos una pos
            Call Get_Pos_By_Guild(Userindex, guild_Num, user_toPosX, user_toPosY)

            'tenemos una pos

            If (user_toPosX <> 0) And (user_toPosY <> 0) Then
                Call WarpUserChar(Userindex, MAPA_CVC, user_toPosX, user_toPosY, True)

            End If

            UserList(Userindex).cvcUser.en_CVC = True
        Else
            'no ai mas espacio..
            Call Protocol.WriteConsoleMsg(Userindex, "No puedes entrar al CVC porque tu clan ya tiene " & CStr(CVC_Info.max_Users) & " jugadores.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub
 
Public Function Can_Ingress(ByVal Userindex As Integer, _
                            ByRef errorMsg As String) As Boolean

    '
    ' @ Checkea si puede ingresar.

    Dim guild_Num As Byte

    Can_Ingress = False

    If UserList(Userindex).Stats.ELV < 40 Then
        errorMsg = "No puedes entrar si eres menor del nivel 40."

        Exit Function

    End If

    If UserList(Userindex).GuildIndex = 0 Then
        errorMsg = "No perteneces a ningun clan.!"

        Exit Function

    End If

    If UCase$(modGuilds.GuildLeader(UserList(Userindex).GuildIndex)) <> UCase$(UserList(Userindex).Name) Then
        errorMsg = "No eres el lider de ningun clan."

        Exit Function

    End If

    If UserList(Userindex).cvcUser.en_CVC = True Then
        errorMsg = "Estas en el cvc!"

        Exit Function

    End If

    If UserList(Userindex).flags.Muerto <> 0 Then
        errorMsg = "Estas muerto."

        Exit Function

    End If

    If UserList(Userindex).Counters.Pena <> 0 Then
        errorMsg = "Estas en la carcel."

        Exit Function

    End If

    'obtengo el 1 o 2 segun su gi
    guild_Num = Find_Guild_Num(UserList(Userindex).GuildIndex)

    'tenemos clan

    If (guild_Num <> 0) Then

        'llego al tope de usuario?

        If index_In_CVC(guild_Num) = 0 Then
            errorMsg = "El clan llego al limite de usuarios en este clan vs clan."

            Exit Function

        End If

    Else
        errorMsg = "Tu clan no esta en ningun CVC."

        Exit Function

    End If

    Can_Ingress = True

End Function
 
Private Function index_In_CVC(ByVal guild_Num As Byte) As Byte

    '
    ' @ Devuelve un slot para el array de usuarios de cada clan.

    Dim j As Long

    With CVC_Info.Guild(guild_Num)

        For j = 1 To UBound(.UsUaRiOs())

            'slot vacio?

            If .UsUaRiOs(j) = -1 Then
                index_In_CVC = CByte(j)

                Exit Function

            Else   'hay un user?

                'no tiene un id valido?

                If UserList(.UsUaRiOs(j)).ConnID = -1 Then
                    index_In_CVC = CByte(j)

                    Exit Function

                End If

            End If

        Next j

    End With

    index_In_CVC = 0

End Function
 
Private Sub Get_Pos_By_Guild(ByVal Userindex As Integer, _
                             ByVal guild_Num As Byte, _
                             ByRef tPosX As Byte, _
                             ByRef tPosY As Byte)

    '
    ' @ Devuelve una pos para un usuario.

    If (guild_Num = 1) Then
        tPosX = PRIMER_CLAN_X
        tPosY = PRIMER_CLAN_Y
    Else
        tPosX = SECOND_CLAN_X
        tPosY = SECOND_CLAN_Y

    End If

    'tenemos la pos default , ahora buscamos un tile
    Call FindLegalPos(Userindex, MAPA_CVC, CInt(tPosX), CInt(tPosY))

End Sub
 
Private Function Find_Guild_Num(ByVal Guild_Index As Integer) As Byte

    '
    ' @ Devuelve 1 o 2 segun el guildIndex

    With CVC_Info

        If .Guild(1).Guild_Index = Guild_Index Then
            Find_Guild_Num = 1
        ElseIf .Guild(2).Guild_Index = Guild_Index Then
            Find_Guild_Num = 2

        End If

    End With

End Function
 
Private Function Get_Num_Dies(ByVal guild_Num As Byte) As Byte

    '
    ' @ Devuelve la cantidad de muertos.

    With CVC_Info

        Dim j As Long

        For j = 1 To UBound(.Guild(guild_Num).UsUaRiOs())

            With .Guild(guild_Num)

                'tenemos un ui

                If (.UsUaRiOs(j) <> -1) Then

                    'si no esta logeado o esta muerto.

                    If (UserList(.UsUaRiOs(j)).ConnID = -1) Or (UserList(.UsUaRiOs(j)).flags.Muerto <> 0) Then
                        Get_Num_Dies = Get_Num_Dies + 1

                    End If

                Else ' no ai ui
                    Get_Num_Dies = Get_Num_Dies + 1

                End If

            End With

        Next j

    End With

End Function

