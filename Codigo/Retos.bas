Attribute VB_Name = "Retos"
' Lautaro Leonel Marino. Luján, Buenos Aires
' 13/08/2019
' Módulo de retos.
' Desde acá se manejan los retos desde 1vs1 hasta nVSn, configurable de forma fácil.

Option Explicit


Private Const MAX_RETOS_SIMULTANEOS As Byte = 4

Public Enum eTipoReto
    None = 0
    FightOne = 1
    FightTwo = 2
    FightThree = 3
End Enum

Public Type tRetoUser
    UserIndex As Integer
    Team As Byte
    Rounds As Byte
End Type

Private Type tMapEvent
    Map As Integer
    x As Byte
    y As Byte
    X2 As Byte
    Y2 As Byte
End Type

Private Type tRetos
    Run As Boolean
    Users() As tRetoUser
    RequiredGld As Long
End Type

Public Arenas(1 To MAX_RETOS_SIMULTANEOS) As tMapEvent
Public Retos(1 To MAX_RETOS_SIMULTANEOS) As tRetos

Public Sub LoadArenas()
10        Arenas(1).Map = 1
20        Arenas(1).x = 25
30        Arenas(1).X2 = 40
40        Arenas(1).y = 19
50        Arenas(1).Y2 = 32
          
60        Arenas(2).Map = 1
70        Arenas(2).x = 61
80        Arenas(2).X2 = 76
90        Arenas(2).y = 19
100       Arenas(2).Y2 = 32
          
110       Arenas(3).Map = 1
120       Arenas(3).x = 25
130       Arenas(3).X2 = 40
140       Arenas(3).y = 46
150       Arenas(3).Y2 = 59
          
160       Arenas(4).Map = 1
170       Arenas(4).x = 59
180       Arenas(4).X2 = 74
190       Arenas(4).y = 46
200       Arenas(4).Y2 = 59

End Sub
Private Sub ResetDueloUser(ByVal UserIndex As Integer)

10    On Error GoTo error

20        With UserList(UserIndex)
30            If .Counters.TimeFight > 0 Then
40                .Counters.TimeFight = 0
50                WriteUserInEvent UserIndex
60            End If
              
70            With Retos(.flags.SlotReto)
80                .Users(UserList(UserIndex).flags.SlotRetoUser).UserIndex = 0
90                .Users(UserList(UserIndex).flags.SlotRetoUser).Team = 0
100               .Users(UserList(UserIndex).flags.SlotRetoUser).Rounds = 0
110           End With
              
120           .flags.SlotReto = 0
130           .flags.SlotRetoUser = 255
140           StatsDuelos UserIndex
150           WarpPosAnt UserIndex
160       End With
          
170   Exit Sub

error:
180
End Sub
Private Sub ResetDuelo(ByVal SlotReto As Byte)
10        On Error GoTo error

          Dim LoopC As Integer
          
20        With Retos(SlotReto)
30            For LoopC = LBound(.Users()) To UBound(.Users())
              
40                If .Users(LoopC).UserIndex > 0 Then
50                    ResetDueloUser .Users(LoopC).UserIndex
60                End If
                  
70                .Users(LoopC).UserIndex = 0
80                .Users(LoopC).Rounds = 0
90                .Users(LoopC).Team = 0

100           Next LoopC
          
130           .RequiredGld = 0
140           .Run = False
150       End With
          
160   Exit Sub

error:
170       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetDuelo()"
End Sub
Private Function FreeSlotArena() As Byte
          Dim LoopC As Integer
          
10        For LoopC = 1 To MAX_RETOS_SIMULTANEOS
20            If Retos(LoopC).Run = False Then
30                FreeSlotArena = LoopC
40                Exit Function
50            End If
60        Next LoopC
End Function
Private Function FreeSlot() As Byte
          ' • Slot libre para comenzar un nuevo enfrentamiento
          Dim LoopC As Integer
          
10        FreeSlot = 0
          
20        For LoopC = 1 To MAX_RETOS_SIMULTANEOS
30            With Retos(LoopC)
40                If .Run = False Then
50                    FreeSlot = LoopC
60                    Exit For
70                End If
80            End With
90        Next LoopC
          
End Function

Private Sub PasateInteger(ByVal SlotArena As Byte, ByRef Users() As String)
10        On Error GoTo error

          ' Cuando se acepta un reto los UserId strings pasan a UserId integer
          
20        With Retos(SlotArena)
              Dim LoopC As Integer
              
30            ReDim .Users(LBound(Users()) To UBound(Users())) As tRetoUser
              
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                .Users(LoopC).UserIndex = NameIndex(Users(LoopC))
                  
60                If .Users(LoopC).UserIndex > 0 Then
80                    UserList(.Users(LoopC).UserIndex).Stats.Gld = UserList(.Users(LoopC).UserIndex).Stats.Gld - .RequiredGld
90                    WriteUpdateGold .Users(LoopC).UserIndex
100               End If
                  
110           Next LoopC
120       End With
130   Exit Sub

error:
140       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : PasateInteger()"
End Sub

Private Sub RewardUsers(ByVal SlotReto As Byte, ByVal UserIndex As Integer)
10        On Error GoTo error
          
          Dim Obj As Obj
          
20        With UserList(UserIndex)
50            .Stats.Gld = .Stats.Gld + (Retos(SlotReto).RequiredGld * 2)
60            WriteUpdateGold UserIndex
120       End With
          
130   Exit Sub

error:
140       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : RewardUsers()"
End Sub
Private Function SetSubTipo(ByRef Users() As String) As eTipoReto
10        On Error GoTo error
          
20        If UBound(Users()) = 1 Then
30            SetSubTipo = FightOne
40            Exit Function
50        End If
          
60        If UBound(Users()) = 3 Then
70            SetSubTipo = FightTwo
80            Exit Function
90        End If
          
100       If UBound(Users()) = 5 Then
110           SetSubTipo = FightThree
120           Exit Function
130       End If
          
140       SetSubTipo = 0
150   Exit Function

error:
160       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetSubTipo()"
End Function
Private Function CanSetUsers(ByRef Users() As String) As Boolean
10        On Error GoTo error
          
          Dim tUser As Integer
          Dim tmpUsers() As String
          
          Dim LoopC As Integer, loopX As Integer
          Dim Tmp As String
          
          ' • Chequeos de cantidad de personajes teniendo en cuenta el tipo de reto.
          
20        If SetSubTipo(Users()) = 0 Then
30            CanSetUsers = False
40            Exit Function
50        End If
          
60        ReDim tmpUsers(LBound(Users()) To UBound(Users())) As String
          
70        For LoopC = LBound(Users()) To UBound(Users())
80            tmpUsers(LoopC) = Users(LoopC)
90        Next LoopC
          
          
100       For LoopC = LBound(Users()) To UBound(Users())
110           For loopX = LBound(Users()) To UBound(Users()) - LoopC
120               If Not loopX = UBound(Users()) Then
130                   If StrComp(UCase$(tmpUsers(loopX)), UCase$(tmpUsers(loopX + 1))) = 0 Then
140                       CanSetUsers = False
150                       Exit Function
160                   Else
170                       Tmp = tmpUsers(loopX)
                          
180                       tmpUsers(loopX) = tmpUsers(loopX + 1)
190                       tmpUsers(loopX + 1) = Tmp
200                   End If
210               End If
220           Next loopX
230       Next LoopC
          
240       CanSetUsers = True
250   Exit Function

error:
260       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanSetUsers()"
End Function

Private Function CanContinueFight(ByVal UserIndex As Integer) As Boolean
10        On Error GoTo error
          
          ' • Si encontramos un personaje vivo el evento continua.
          Dim LoopC As Integer
          Dim SlotReto As Byte
          Dim SlotRetoUser As Byte
          
20        SlotReto = UserList(UserIndex).flags.SlotReto
30        SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser

40        CanContinueFight = False
          
50        With Retos(SlotReto)
          
60            For LoopC = LBound(.Users()) To UBound(.Users())
70                If .Users(LoopC).UserIndex > 0 And .Users(LoopC).UserIndex <> UserIndex Then
80                    If .Users(SlotRetoUser).Team = .Users(LoopC).Team Then
90                        With UserList(.Users(LoopC).UserIndex)
100                           If .flags.Muerto = 0 Then
110                               CanContinueFight = True
120                               Exit Function
130                           End If
140                       End With
150                   End If
                      
160               End If
170           Next LoopC
              
180       End With
190   Exit Function

error:
200       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanContinueFight()"
End Function
Private Function AttackerFight(ByVal SlotReto As Byte, ByVal TeamUser As Byte) As Integer
10        On Error GoTo error

          ' • Buscamos al AttackerIndex (Caso abandono del evento)
          Dim LoopC As Integer
          
20        With Retos(SlotReto)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).UserIndex > 0 Then
50                    If .Users(LoopC).Team > 0 And .Users(LoopC).Team <> TeamUser Then
60                        AttackerFight = .Users(LoopC).UserIndex
70                        Exit For
80                    End If
90                End If
100           Next LoopC
110       End With
120   Exit Function

error:
130       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AttackerFight()"
End Function
Private Function CanAcceptFight(ByVal UserIndex As Integer, _
                        ByVal UserName As String) As Boolean

10        On Error GoTo error
          
          ' • Username es el que mando el reto al principio.
          ' • Si está online y cumple con los requisitos entra
          Dim SlotTemp As Byte
          Dim tUser As Integer
          Dim ArrayNulo As Long
          
20            tUser = NameIndex(UserName)
              
30            If tUser <= 0 Then
                  ' Personaje offline
40                CanAcceptFight = False
50                Exit Function
60            End If
              
70            With UserList(tUser)
80                'GetSafeArrayPointer .RetoTemp.Users, ArrayNulo
90                'If ArrayNulo <= 0 Then Exit Function
                  
100               SlotTemp = SearchFight(UCase$(UserList(UserIndex).Name), .RetoTemp.Users, .RetoTemp.Accepts)
                  
110               If SlotTemp = 255 Then
120                   CanAcceptFight = False
                      ' El personaje no te mando ninguna solicitud
130                   Exit Function
140               End If
                  
150               If .RetoTemp.Accepts(SlotTemp) = 1 Then
                      ' El personaje ya aceptó.
160                   CanAcceptFight = False
170                   Exit Function
180               End If
                  
                  
                  ' Valido el usuario
190               .RetoTemp.Accepts(SlotTemp) = 1
200               CanAcceptFight = True
                  
                  ' • Chequeo de aceptaciones
210               If CheckAccepts(.RetoTemp.Accepts) Then
220                   GoFight tUser
230               End If
          
          
240           End With
              
250   Exit Function

error:
260       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAcceptFight()"
End Function
Private Function ValidateFight_Users(ByVal UserIndex As Integer, _
                                    ByVal GldRequired As Long, _
                                    ByRef Users() As String) As Boolean
                                              
10        On Error GoTo error
          
          ' • Validamos al Team seleccionado.
          
          Dim LoopC As Integer
          Dim tUser As Integer
                                     
20        For LoopC = LBound(Users()) To UBound(Users())
30            If Users(LoopC) <> vbNullString Then
40                tUser = NameIndex(Users(LoopC))
                  
                  ' No fuckings gms
                  If tUser > 0 Then
                      If EsGm(tUser) Then
                         ' ValidateFight_Users = False
                         ' Exit Function
                      End If
                  End If
                  
50                If tUser <= 0 Then
60                    'SendMsjUsers "El personaje " & Users(LoopC) & " está offline.", Users()
                      WriteConsoleMsg UserIndex, "El personaje " & Users(LoopC) & " está offline", FontTypeNames.FONTTYPE_INFO
70                    ValidateFight_Users = False
80                    Exit Function
90                End If
                  
100               With UserList(tUser)
110                   If .flags.Muerto = 1 Then
                          WriteConsoleMsg UserIndex, "El personaje " & Users(LoopC) & " está muerto.", FontTypeNames.FONTTYPE_INFO
130                       ValidateFight_Users = False
140                       Exit Function
150                   End If
                      
160                   If MapInfo(.Pos.Map).Pk = True Then
180                       ValidateFight_Users = False
190                       Exit Function
200                   End If
                      
210                   If (.flags.SlotReto > 0) Then
                          WriteConsoleMsg UserIndex, "El personaje " & Users(LoopC) & " está participando en otro evento.", FontTypeNames.FONTTYPE_INFO
230                       ValidateFight_Users = False
240                       Exit Function
250                   End If
                      
260                   If .flags.Comerciando Then
                          WriteConsoleMsg UserIndex, "El personaje " & Users(LoopC) & " no está disponible en este momento.", FontTypeNames.FONTTYPE_INFO
280                       ValidateFight_Users = False
290                       Exit Function
300                   End If
                      
380                   If .Stats.Gld < GldRequired Then
                          WriteConsoleMsg UserIndex, "El personaje " & .Name & " no tiene las monedas en su billetera.", FontTypeNames.FONTTYPE_INFO
400                       ValidateFight_Users = False
410                       Exit Function
420                   End If

500               End With
510           End If
520       Next LoopC
          
          
530       ValidateFight_Users = True
          
540   Exit Function

error:
550       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight_Users()"
End Function
Private Function ValidateFight(ByVal UserIndex As Integer, _
                                ByVal GldRequired As Long, _
                                ByRef Users() As String) As Boolean
                                      
10        On Error GoTo error
          
              ' • Validamos el enfrentamiento que se va a disputar
              ' • UserIndex = Personaje que inició la invitación.
              '(Userindex, Tipo, GldRequired, Users) Then
              
          Dim LoopC As Integer
          Dim tUser As Integer

70        If GldRequired < 0 Or GldRequired > 100000000 Then
              WriteConsoleMsg UserIndex, "Oro Mínimo: 0 . Oro Máximo 100.000.000", FontTypeNames.FONTTYPE_INFO
90            ValidateFight = False
100           Exit Function
110       End If
          
          ' • Los Team están diferentes en cuanto a cantidad. [LOG ERROR ANTI CHEAT]
120       If Not CanSetUsers(Users) Then
              'Mensaje: Intento hackear el sistema
130           LogRetos "POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos."
140           ValidateFight = False
150           Exit Function
160       End If
          
          ' Validamos a los personajes
170       If Not ValidateFight_Users(UserIndex, GldRequired, Users()) Then
180           ValidateFight = False
190           Exit Function
200       End If
          
          
210       ValidateFight = True
          
220   Exit Function

error:
230       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight()"
End Function

Private Function StrTeam(ByRef Users() As tRetoUser) As String
          
10        On Error GoTo error
          
          ' • Devuelve ENEMIGOS vs TEAM
          
          Dim LoopC As Integer
          Dim strtemp(1) As String
          
          ' 1 vs 1
20        If UBound(Users()) = 1 Then
30            If Users(0).UserIndex > 0 Then
40                strtemp(0) = UserList(Users(0).UserIndex).Name
50            Else
60                strtemp(0) = "Usuario descalificado"
70            End If
              
80            If Users(1).UserIndex > 0 Then
90                strtemp(1) = UserList(Users(1).UserIndex).Name
100           Else
110               strtemp(1) = "Usuario descalificado"
120           End If
              
130           StrTeam = strtemp(0) & " vs " & strtemp(1)
140           Exit Function
150       End If
          
160       For LoopC = LBound(Users()) To UBound(Users())
170           If Users(LoopC).UserIndex > 0 Then
180               If LoopC < ((1 + UBound(Users)) / 2) Then
190                   strtemp(0) = strtemp(0) & UserList(Users(LoopC).UserIndex).Name & ", "
200               Else
210                   strtemp(1) = strtemp(1) & UserList(Users(LoopC).UserIndex).Name & ", "
220               End If
230           End If
240       Next LoopC
          
250       If Not strtemp(0) = vbNullString Then
260           strtemp(0) = Left$(strtemp(0), Len(strtemp(0)) - 2)
270       Else
280           strtemp(0) = "Equipo descalificado"
290       End If
          
300       If Not strtemp(1) = vbNullString Then
310           strtemp(1) = Left$(strtemp(1), Len(strtemp(1)) - 2)
320       Else
330           strtemp(1) = "Equipo descalificado"
340       End If
          
350       StrTeam = strtemp(0) & " vs " & strtemp(1)
          
360   Exit Function

error:
370       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StrTeam()"
End Function

Private Function CheckAccepts(ByRef Accepts() As Byte) As Boolean
10        On Error GoTo error
          
          ' • Si encontramos a un usuario que no haya aceptado retornamos false.
          Dim LoopC As Integer
          
20        CheckAccepts = True
          
30        For LoopC = LBound(Accepts()) To UBound(Accepts())
40            If Accepts(LoopC) = 0 Then
50                CheckAccepts = False
60                Exit Function
70            End If
80        Next LoopC
          
90    Exit Function

error:
100       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CheckAccepts()"
End Function

Private Function SearchFight(ByVal UserName As String, _
                                ByRef Users() As String, _
                                ByRef Accepts() As Byte) As Byte
                                      
          ' • Buscamos la invitación que nos realizo el personaje UserName
          
10    On Error GoTo error

          Dim LoopC As Integer
          
20        SearchFight = 255
          
30        For LoopC = LBound(Users()) To UBound(Users())
40            If StrComp(Users(LoopC), UserName) = 0 And Accepts(LoopC) = 0 Then
50                    SearchFight = LoopC
60                Exit Function
70            End If
80        Next LoopC
          
90    Exit Function

error:
100       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SearchFight()"
End Function
Public Function CanAttackReto(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
          
10    On Error GoTo error

20        CanAttackReto = True
          
30        With UserList(AttackerIndex)
40            If .flags.SlotReto > 0 Then
                  
                  'If Retos(.flags.SlotReto).Users(.flags.SlotRetoUser).Team = _
                      Retos(.flags.SlotReto).Users(UserList(VictimIndex).flags.SlotRetoUser).Team Then
50                    CanAttackReto = True
60                    Exit Function
                  'End If
70            End If
          
80        End With
          
90    Exit Function

error:
100       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackReto()"
End Function

Private Sub SendInvitation(ByVal UserIndex As Integer, _
                            ByVal GldRequired As Long, _
                            ByRef Users() As String)
                                  
10        On Error GoTo error
          
          ' • Enviamos la solicitud del duelo a los demás y guardamos los datos temporales al usuario mandatario.
          
          Dim LoopC As Integer
          Dim strtemp As String
          Dim tUser As Integer
          Dim str() As tRetoUser
          
          ' • Save data temp
20        With UserList(UserIndex)
          
              
30            With .RetoTemp
40                ReDim .Accepts(LBound(Users()) To UBound(Users())) As Byte
50                ReDim .Users(LBound(Users()) To UBound(Users())) As String
                  
60                .RequiredGld = GldRequired
90                .Users = Users
                  
110               .Accepts(UBound(Users())) = 1 ' El último personaje es el que envió por lo tanto ya aceptó.
120           End With
130       End With
          
140       ReDim str(LBound(Users()) To UBound(Users())) As tRetoUser
          
150       For LoopC = LBound(Users()) To UBound(Users())
160           str(LoopC).UserIndex = NameIndex(Users(LoopC))
170       Next LoopC
          
180       strtemp = StrTeam(str) & "."
200       strtemp = strtemp & IIf(GldRequired > 0, " Oro requerido: " & GldRequired & ".", vbNullString)
220       strtemp = strtemp & " Para aceptar tipea /ACEPTAR " & UserList(UserIndex).Name
          
230       For LoopC = LBound(Users()) To UBound(Users())
240           tUser = NameIndex(Users(LoopC))
              
250           If tUser <> UserIndex Then
260               WriteConsoleMsg tUser, strtemp, FontTypeNames.FONTTYPE_INFO
270           End If
                                              
280       Next LoopC
          
290   Exit Sub

error:
300       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendInvitation()"
End Sub



Private Sub GoFight(ByVal UserIndex As Integer)
          ' • Comienzo del duelo
          
10    On Error GoTo error

          Dim GldRequired As Long
          Dim SlotArena As Byte
          
20        SlotArena = FreeSlotArena
          
30        If SlotArena = 0 Then
              ' Mensaje : No hay mas arenas disponibles
40            Exit Sub
50        End If
          
60        With UserList(UserIndex)
70            If ValidateFight(UserIndex, .RetoTemp.RequiredGld, .RetoTemp.Users) Then
                  
100               Retos(SlotArena).RequiredGld = .RetoTemp.RequiredGld
110               Retos(SlotArena).Run = True
                  
120               PasateInteger SlotArena, .RetoTemp.Users
                  
130               SetUserEvent SlotArena, Retos(SlotArena).Users
140               WarpFight Retos(SlotArena).Users
150           End If
160       End With
          
170   Exit Sub

error:
180       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : GoFight()"
End Sub
Private Sub SetUserEvent(ByVal SlotReto As Byte, ByRef Users() As tRetoUser)

10        On Error GoTo error
          ' • Guardamos los slot en los usuarios y seteamos el team.
          
          Dim LoopC As Integer
          Dim SlotRetoUser As Byte
          
20        For LoopC = LBound(Users()) To UBound(Users())
30            If Users(LoopC).UserIndex > 0 Then
40                With Users(LoopC)
50                    If .UserIndex > 0 Then
60                        UserList(.UserIndex).flags.SlotReto = SlotReto
70                        UserList(.UserIndex).flags.SlotRetoUser = LoopC
                          
80                    End If
90                End With
              
100               With Retos(SlotReto)
110                   If LoopC < ((1 + UBound(Users())) / 2) Then
120                       .Users(LoopC).Team = 2
130                   Else
140                       .Users(LoopC).Team = 1
150                   End If
160               End With
              
170               With UserList(Users(LoopC).UserIndex)
180                   .PosAnt.Map = .Pos.Map
190                   .PosAnt.x = .Pos.x
200                   .PosAnt.y = .Pos.y
                      
210               End With
220           End If
230       Next LoopC
          
240   Exit Sub

error:
250       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetUserEvent()"
End Sub
Private Sub WarpFight(ByRef Users() As tRetoUser)

          ' • Teletransportamos a los personajes a la sala de combate
          
10    On Error GoTo error

          Dim LoopC As Integer
          Dim tUser As Integer
          Dim Pos As WorldPos
          Const Tile_Extra As Byte = 5
          
20        For LoopC = LBound(Users()) To UBound(Users())
30            tUser = Users(LoopC).UserIndex
              
40            If tUser > 0 Then
50                Pos.Map = Arenas(UserList(tUser).flags.SlotReto).Map
                  
60                If Users(LoopC).Team = 1 Then
70                    Pos.x = Arenas(UserList(tUser).flags.SlotReto).x
80                    Pos.y = Arenas(UserList(tUser).flags.SlotReto).y
90                Else
100                   Pos.x = Arenas(UserList(tUser).flags.SlotReto).X2
110                   Pos.y = Arenas(UserList(tUser).flags.SlotReto).Y2
120               End If
                  
130               With UserList(tUser)
140                   .Counters.TimeFight = 10
150                   Call WriteUserInEvent(tUser)
                      ' Mensaje: ¡Preparate en 10 segundos comenzarás a luchar!
                  
160                   ClosestStablePos Pos, Pos
170                   WarpUserChar tUser, Pos.Map, Pos.x, Pos.y, False
180               End With
190           End If
200       Next LoopC
          
210   Exit Sub

error:
220       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : WarpFight()"
End Sub

Private Sub AddRound(ByVal SlotReto As Byte, ByVal Team As Byte)

10    On Error GoTo error

          Dim LoopC As Integer
20        With Retos(SlotReto)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Team = Team And .Users(LoopC).UserIndex > 0 Then
50                    .Users(LoopC).Rounds = .Users(LoopC).Rounds + 1
60                End If
70            Next LoopC
          
80        End With
          
90    Exit Sub

error:
100       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AddRound()"
End Sub
Private Sub SendMsjUsers(ByVal strMsj As String, _
                        ByRef Users() As String)
                              
10    On Error GoTo error

          Dim LoopC As Integer
          Dim tUser As Integer
          
20        For LoopC = LBound(Users()) To UBound(Users())
30            tUser = NameIndex(Users(LoopC))
40            If tUser > 0 Then
50                WriteConsoleMsg tUser, strMsj, FontTypeNames.FONTTYPE_VENENO
60            End If
70        Next LoopC
          
80    Exit Sub

error:
90        LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendMsjUsers()"
End Sub

Private Function ExistCompañero(ByVal UserIndex As Integer) As Boolean
          Dim LoopC As Integer
          Dim SlotReto As Byte
          Dim SlotRetoUser As Byte
          
   On Error GoTo ExistCompañero_Error

10        SlotReto = UserList(UserIndex).flags.SlotReto
20        SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser
          
30        With Retos(SlotReto)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).UserIndex > 0 Then
60                    If LoopC <> SlotRetoUser Then
70                        If .Users(LoopC).Team = .Users(SlotRetoUser).Team Then
80                            ExistCompañero = True
90                            Exit For
100                       End If
110                   End If
120               End If
130           Next LoopC
140       End With

   On Error GoTo 0
   Exit Function

ExistCompañero_Error:

    LogRetos "Error " & Err.Number & " (" & Err.description & ") in procedure ExistCompañero of Módulo mRetos in line " & Erl
          
End Function
Public Sub UserdieFight(ByVal UserIndex As Integer, ByVal AttackerIndex As Integer, ByVal Forzado As Boolean)

10    On Error GoTo error

          ' • Un personaje en reto es matado por otro.
          Dim LoopC As Integer
          Dim strtemp As String
          Dim SlotReto As Byte
          Dim TeamUser As Byte
          Dim Rounds As Byte
          Dim Deslogged As Boolean
          Dim ExistTeam As Boolean
          
20        SlotReto = UserList(UserIndex).flags.SlotReto
          
30        Deslogged = False
              
          ' • Caso hipotetico de deslogeo. El funcionamiento es el mismo, con la diferencia de que se busca al ganador.
40        If AttackerIndex = 0 Then
50            AttackerIndex = AttackerFight(SlotReto, Retos(SlotReto).Users(UserList(UserIndex).flags.SlotRetoUser).Team)
              
60            Deslogged = True
70        End If
          
80        TeamUser = Retos(SlotReto).Users(UserList(AttackerIndex).flags.SlotRetoUser).Team
90        ExistTeam = ExistCompañero(UserIndex)
          
          
          ' Deslogeo de todos los integrantes del team
100       If Forzado Then
110           If Not ExistTeam Then
120               FinishFight SlotReto, TeamUser
130               ResetDuelo SlotReto
140               Exit Sub
150           End If
160       End If
          
170       With UserList(UserIndex)
180           If Not CanContinueFight(UserIndex) Then
190               With Retos(SlotReto)
200                   For LoopC = LBound(.Users()) To UBound(.Users())
210                       With .Users(LoopC)
220                           If .UserIndex > 0 And .Team = TeamUser Then
230                               If Rounds = 0 Then
240                                   AddRound SlotReto, .Team
250                                   Rounds = .Rounds
260                               End If
                                  
270                               WriteConsoleMsg .UserIndex, "Has ganado el round. Rounds ganados: " & .Rounds & ".", FontTypeNames.FONTTYPE_VENENO
                                   
280                           End If
290                       End With
                          
300                       If .Users(LoopC).UserIndex > 0 Then StatsDuelos .Users(LoopC).UserIndex
310                   Next LoopC
                      
320                   If Rounds >= (3 / 2) + 0.5 Or Forzado Then
330                       FinishFight SlotReto, TeamUser
340                       ResetDuelo SlotReto
350                       Exit Sub
360                   Else
370                       FinishFight SlotReto, TeamUser, True
                          'StatsDuelos Userindex
380                   End If
390               End With
400           End If
              
 
410           If Deslogged Then
420               ResetDueloUser UserIndex
430           End If
440       End With
          
450   Exit Sub

error:
460       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : UserdieFight() en linea " & Erl
End Sub


Private Sub StatsDuelos(ByVal UserIndex As Integer)

10    On Error GoTo error

20        With UserList(UserIndex)

            If .flags.Muerto Then
                RevivirUsuario (UserIndex)
                 .Stats.MinHp = .Stats.MaxHp
                 .Stats.MinMAN = .Stats.MaxMAN
                 .Stats.MinSta = .Stats.MaxSta
              
                WriteUpdateUserStats UserIndex
                Exit Sub
            End If
            


            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
              
            WriteUpdateUserStats UserIndex
            
            'If .flags.Paralizado = 1 Then
                '.flags.Paralizado = 0
                'Call WriteParalizeOK(UserIndex)
            'End If
            
100       End With
          
110   Exit Sub

error:
120       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StatsDuelos()"
End Sub

Private Sub FinishFight(ByVal SlotReto As Byte, ByVal Team As Byte, Optional ByVal ChangeTeam As Boolean)

          ' • Finalizamos el reto o el round.
          
10    On Error GoTo error

          Dim LoopC As Integer
          Dim strtemp As String
          
20        With Retos(SlotReto)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).UserIndex > 0 Then
50                    If UserList(.Users(LoopC).UserIndex).Counters.TimeFight > 0 Then
60                        UserList(.Users(LoopC).UserIndex).Counters.TimeFight = 0
70                        WriteUserInEvent .Users(LoopC).UserIndex
80                    End If
                      
90                    If Team = .Users(LoopC).Team Then
100                       If ChangeTeam Then
110                           StatsDuelos .Users(LoopC).UserIndex
120                       Else
130                           .Run = False
140                           StatsDuelos .Users(LoopC).UserIndex
150                           RewardUsers SlotReto, .Users(LoopC).UserIndex
                              
160                           If .Users(LoopC).Rounds > 0 Then
170                               WriteConsoleMsg .Users(LoopC).UserIndex, "Has ganado el reto con " & .Users(LoopC).Rounds & " rounds a tu favor.", FontTypeNames.FONTTYPE_VENENO
180                           Else
190                               WriteConsoleMsg .Users(LoopC).UserIndex, "Has ganado el reto.", FontTypeNames.FONTTYPE_VENENO
200                           End If

210                           strtemp = strtemp & UserList(.Users(LoopC).UserIndex).Name & ", "
                              
220                       End If
                      
230                   End If
240               End If
250           Next LoopC
          
260           If ChangeTeam Then
270               Call WarpFight(.Users())
280           Else
290               strtemp = Left$(strtemp, Len(strtemp) - 2)
        
300               SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos» " & StrTeam(.Users()) & ". Ganador " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro", FontTypeNames.FONTTYPE_INFO)
310               LogRetos "Retos» " & StrTeam(.Users()) & ". Ganador el team de " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro"
320           End If
330       End With
          
340   Exit Sub

error:
350       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : FinishFight() en linea " & Erl
End Sub

' • Procedimientos necesarios para enviar,aceptar,abandonar.

Public Sub SendFight(ByVal UserIndex As Integer, _
                            ByVal GldRequired As Long, _
                            ByRef Users() As String)
          
10        On Error GoTo error
          
          ' Enviamos una solicitud de enfrentamiento
          
20        With UserList(UserIndex)
              
30            If ValidateFight(UserIndex, GldRequired, Users) Then
40                SendInvitation UserIndex, GldRequired, Users
                  
50                WriteConsoleMsg UserIndex, "Espera noticias para concretar el reto que has enviado. Recuerda que si vuelves a mandar, la anterior solicitud se cancela.", FontTypeNames.FONTTYPE_WARNING
60            End If
              
              
70        End With
          
80    Exit Sub
error:
90        LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendFight()"
End Sub

Public Sub AcceptFight(ByVal UserIndex As Integer, _
                        ByVal UserName As String)
                              
10    On Error GoTo error
                              
20        With UserList(UserIndex)
              
30            If CanAcceptFight(UserIndex, UserName) Then
                  
40                WriteConsoleMsg UserIndex, "Has aceptado la invitación.", FontTypeNames.FONTTYPE_INFO
                  ' Has aceptado la invitacion bababa
50            End If
60        End With
          
70    Exit Sub
error:
80        LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AcceptFight()"
End Sub

Public Sub WarpPosAnt(ByVal UserIndex As Integer)
          ' • Warpeo del personaje a su posición anterior.
          
          Dim Pos As WorldPos
          
   On Error GoTo WarpPosAnt_Error

10        With UserList(UserIndex)
20            Pos.Map = .PosAnt.Map
30            Pos.x = .PosAnt.x
40            Pos.y = .PosAnt.y
                          
50            Call FindLegalPos(UserIndex, Pos.Map, Pos.x, Pos.y)
60            Call WarpUserChar(UserIndex, Pos.Map, Pos.x, Pos.y, False)
              
70            .PosAnt.Map = 0
80            .PosAnt.x = 0
90            .PosAnt.y = 0
          
100       End With

   On Error GoTo 0
   Exit Sub

WarpPosAnt_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure WarpPosAnt of Módulo General in line " & Erl
End Sub

