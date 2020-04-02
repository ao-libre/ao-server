Attribute VB_Name = "mMainLoop"
Option Explicit

Public Type tMainLoop

    MAXINT As Long
    LastCheck As Long

End Type
 
Private Const NumTimers          As Byte = 4

Public MainLoops(1 To NumTimers) As tMainLoop

Public Enum eTimers

    eGameTimer = 1      'Stats entre otros
    epacketResend = 2   'Socket
    eAuditoria = 3      'Centinela
    TimerAI = 4         'Npc's

End Enum

Public prgRun As Boolean
 
Public Sub MainLoop()
    
    Dim LoopC As Integer

    MainLoops(eTimers.eGameTimer).MAXINT = 40
    MainLoops(eTimers.epacketResend).MAXINT = 10
    MainLoops(eTimers.eAuditoria).MAXINT = 1000
    MainLoops(eTimers.TimerAI).MAXINT = 380
    
    prgRun = True
    
    Do While prgRun
        frmMain.lblLloviendoInfo.Caption = "Esta lloviendo? - " & IIf(Lloviendo, "Si, usa paraguas", "No, totalmente soleado")

        For LoopC = 1 To NumTimers

            If GetTickCount - MainLoops(LoopC).LastCheck >= MainLoops(LoopC).MAXINT Then
                Call MakeProcces(LoopC)

            End If

            DoEvents
        Next LoopC

    Loop
    
End Sub
 
Private Sub MakeProcces(ByVal index As Integer)
    
    Select Case index
    
        Case eTimers.eGameTimer
            Call GameTimer
 
        Case eTimers.epacketResend
            Call packetResend
            
        Case eTimers.eAuditoria
            Call Auditoria
            
        Case eTimers.TimerAI
            Call TIMER_AI
            
    End Select
    
    MainLoops(index).LastCheck = GetTickCount
    
End Sub

Private Sub Auditoria()

    On Error GoTo errhand
    
    Call PasarSegundo 'sistema de desconexion de 10 segs
    
    Static centinelSecs As Byte

    centinelSecs = centinelSecs + 1

    If centinelSecs = 5 Then
        'Every 5 seconds, we try to call the player's attention so it will report the code.
        Call modCentinela.AvisarUsuarios
    
        centinelSecs = 0

    End If

    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)

End Sub

Private Sub packetResend()

    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Attempts to resend to the user all data that may be enqueued.
    '***************************************************
    On Error GoTo ErrHandler:

    Dim i As Long
    
    For i = 1 To LastUser

        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.Length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.Length))

            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)

    Resume Next

End Sub

Private Sub TIMER_AI()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Long

    Dim mapa     As Integer

    Dim e_p      As Integer
    
    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then

        'Update NPCs
        For NpcIndex = 1 To LastNPC
            
            With Npclist(NpcIndex)

                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                
                    ' Chequea si contiua teniendo dueno
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else

                        ' Preto? Tienen ai especial
                        If .NPCtype = eNPCType.Pretoriano Then
                            Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                        Else

                            'Usamos AI si hay algun user en el mapa
                            If .flags.Inmovilizado = 1 Then
                                Call EfectoParalisisNpc(NpcIndex)

                            End If
                            
                            mapa = .Pos.Map
                            
                            If mapa > 0 Then
                                If MapInfo(mapa).NumUsers > 0 Then
                                    If .Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End With

        Next NpcIndex

    End If
    
    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub GameTimer()

    '********************************************************
    'Author: Unknown
    'Last Modify Date: -
    '********************************************************
    Dim iUserIndex   As Long

    Dim bEnviarStats As Boolean

    Dim bEnviarAyS   As Boolean
    
    On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then
                'User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                        
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)

                        End If
                        
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo Then
                                If Not Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                        'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                    Else
                                        'esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        'termina de descansar automaticamente
                                        If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False

                                        End If
                                        
                                    End If

                                End If

                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False

                                    End If
                                    
                                Else
                                    'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    'termina de descansar automaticamente
                                    If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False

                                    End If
                                    
                                End If

                            End If

                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else

                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                    End If 'Muerto

                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1

                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)

                    End If

                End If 'UserLogged
                
                'Si quedo algo para mandar, lo mandamos.
                Call FlushBuffer(iUserIndex)
                
                'Ya terminamos de procesar el paquete, sigamos recibiendo.
                .Counters.PacketsTick = 0
                
            End If

        End With

    Next iUserIndex

    Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)

End Sub

Public Sub PasarSegundo()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim i As Long
    
    'Limpieza del mundo
    If counterSV.Limpieza > 0 Then
        counterSV.Limpieza = counterSV.Limpieza - 1
        
        If counterSV.Limpieza < 6 Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo en " & counterSV.Limpieza & " segundos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
        
        If counterSV.Limpieza = 0 Then
            Call BorrarObjetosLimpieza
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo finalizada.", FontTypeNames.FONTTYPE_SERVER))
            UltimoSlotLimpieza = -1

        End If

    End If
    
    For i = 1 To LastUser

        With UserList(i)

            If .flags.UserLogged Then

                'Cerrar usuario
                If .Counters.Saliendo Then
                    .Counters.Salir = .Counters.Salir - 1

                    If .Counters.Salir <= 0 Then
                        Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call CloseSocket(i, True)
                        Call FlushBuffer(i)
                        Call LoginAccountCharfile(i, UserList(i).mail)
                        Call FlushBuffer(i)

                    End If

                End If
            
                ' Conteo de los Retos
                If .Counters.TimeFight > 0 Then
                    .Counters.TimeFight = .Counters.TimeFight - 1
                    
                    ' Cuenta regresiva de retos y eventos
                    If .Counters.TimeFight = 0 Then
                        WriteConsoleMsg i, "Cuenta -> YA!", FontTypeNames.FONTTYPE_FIGHT
                                             
                        If .flags.SlotReto > 0 Then
                            Call WriteUserInEvent(i)
                        End If
                    Else
                        WriteConsoleMsg i, "Cuenta -> " & .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD
                    End If
                End If
                
                If .Counters.Pena > 0 Then

                    'Restamos las penas del personaje
                    If .Counters.Pena > 0 Then
                        .Counters.Pena = .Counters.Pena - 1
                 
                        If .Counters.Pena < 1 Then
                            .Counters.Pena = 0
                            Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                            Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                            Call FlushBuffer(i)

                        End If

                    End If
                    
                End If
                
                'Sacamos energia
                If Lloviendo Then Call EfectoLluvia(i)
                
                If Not .Pos.Map = 0 Then

                    'Counter de piquete
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                            If .flags.Muerto = 0 Then
                                .Counters.PiqueteC = .Counters.PiqueteC + 1
                                .Counters.ContadorPiquete = .Counters.ContadorPiquete + 1
                                If .Counters.ContadorPiquete = 6 Then
                                    Call WriteConsoleMsg(i, "Estas obstruyendo la via publica, muevete o seras encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                                    .Counters.ContadorPiquete = 0
                                End If
                                If .Counters.PiqueteC >= 30 Then
                                    .Counters.PiqueteC = 0
                                    .Counters.ContadorPiquete = 0
                                    Call Encarcelar(i, MinutosCarcelPiquete)
                                End If
                                
                            Call FlushBuffer(i)
                        Else
                            .Counters.PiqueteC = 0

                        End If

                    Else
                        .Counters.PiqueteC = 0

                    End If

                End If

            End If

        End With

    Next i

    Exit Sub

ErrHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub
