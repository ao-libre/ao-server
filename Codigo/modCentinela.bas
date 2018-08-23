Attribute VB_Name = "modCentinela"
'*****************************************************************
'modCentinela.bas - ImperiumAO - v1.2
'
'Funciónes de control para usuarios que se encuentran trabajando
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
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

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   ImperiumAO 1.2
'   - First Relase
'
'Juan Martín Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.11.5
'   - Small improvements and added logs to detect possible cheaters
'
'Juan Martín Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.12.0
'   - Added several messages to spam users until they reply
'
'ZaMa
'   Alkon AO 0.13.0
'   - Added several paralel checks
'*****************************************************************

Option Explicit

Private Const NPC_CENTINELA As Integer = 16  'Índice del NPC en el .dat

Private Const TIEMPO_INICIAL As Byte = 2 'Tiempo inicial en minutos. No reducir sin antes revisar el timer que maneja estos datos.
Private Const TIEMPO_PASAR_BASE As Integer = 20 'Tiempo minimo fijo para volver a pasar
Private Const TIEMPO_PASAR_RANDOM As Integer = 10 'Tiempo máximo para el random para que el centinela vuelva a pasar

Private Type tCentinela
    NpcIndex As Integer             ' Index of centinela en el servidor
    RevisandoUserIndex As Integer   '¿Qué índice revisamos?
    TiempoRestante As Integer       '¿Cuántos minutos le quedan al usuario?
    clave As Integer                'Clave que debe escribir
    SpawnTime As Long
    Activo As Boolean
End Type

Public centinelaActivado As Boolean

'Guardo cuando voy a resetear a la lista de usuarios del centinela
Private centinelaStartTime As Long
Private centinelaInterval As Long

Private DetenerAsignacion As Boolean

Public Const NRO_CENTINELA As Byte = 5
Public Centinela(1 To NRO_CENTINELA) As tCentinela

Public Sub CallUserAttention()
'*************************************************
'Author: Unknown
'Last modified: 03/10/2010
'Makes noise and FX to call the user's attention.
'03/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************

    'Esta el sistema habilitado?
    If Not centinelaActivado Then Exit Sub

    Dim index As Integer
    Dim UserIndex As Integer
    
    Dim TActual As Long
    TActual = (GetTickCount() And &H7FFFFFFF)
    
    ' Chequea todos los centinelas
    For index = 1 To NRO_CENTINELA
        
        With Centinela(index)
            
            ' Centinela activo?
            If .Activo Then
            
                UserIndex = .RevisandoUserIndex
                
                ' Esta revisando un usuario?
                If UserIndex <> 0 Then
                    
                    If TActual - .SpawnTime >= 5000 Then
                    
                        If Not UserList(UserIndex).flags.CentinelaOK Then
                            Call WritePlayWave(UserIndex, SND_WARP, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y)
                            Call WriteCreateFX(UserIndex, Npclist(.NpcIndex).Char.CharIndex, FXIDs.FXWARP, 0)
                            
                            'Resend the key
                            Call CentinelaSendClave(UserIndex, index)
                            
                            Call FlushBuffer(UserIndex)
                        End If
                    End If
                End If
            End If
        End With
        
    Next index
End Sub

Private Sub GoToNextWorkingChar()
'*************************************************
'Author: Unknown
'Last modified: 03/10/2010
'Va al siguiente usuario que se encuentre trabajando
'09/27/2010: C4b3z0n - Ahora una vez que termina la lista de usuarios, si se cumplio el tiempo de reset, resetea la info y asigna un nuevo tiempo.
'03/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************

    Dim LoopC As Long
    Dim CentinelaIndex As Integer
    
    CentinelaIndex = GetIdleCentinela(1)
    
    For LoopC = 1 To LastUser
        
        With UserList(LoopC)
            
            ' Usuario trabajando y no revisado?
            If .flags.UserLogged And .Counters.Trabajando > 0 And (.flags.Privilegios And PlayerType.User) Then
                If Not .flags.CentinelaOK And .flags.CentinelaIndex = 0 Then
                    'Inicializamos
                    With Centinela(CentinelaIndex)
                        .RevisandoUserIndex = LoopC
                        .TiempoRestante = TIEMPO_INICIAL
                        .clave = RandomNumber(1, 32000)
                        .SpawnTime = GetTickCount() And &H7FFFFFFF
                        .Activo = True
                    
                    
                        'Ponemos al centinela en posición
                        Call WarpCentinela(LoopC, CentinelaIndex)
                        
                        ' Spawneo?
                        If .NpcIndex <> 0 Then
                            'Mandamos el mensaje (el centinela habla y aparece en consola para que no haya dudas)
                            Call WriteChatOverHead(LoopC, "Saludos " & UserList(LoopC).Name & ", soy el Centinela de estas tierras. Me gustaría que escribas /CENTINELA " & .clave & " en no más de dos minutos.", CStr(Npclist(.NpcIndex).Char.CharIndex), vbGreen)
                            Call WriteConsoleMsg(LoopC, "El centinela intenta llamar tu atención. ¡Respóndele rápido!", FontTypeNames.FONTTYPE_CENTINELA)
                            Call FlushBuffer(LoopC)
                            
                            ' Guardo el indice del centinela
                            UserList(LoopC).flags.CentinelaIndex = CentinelaIndex
                        End If
                    
                    End With
                        
                    ' Si ya se asigno un usuario a cada centinela, me voy
                    CentinelaIndex = CentinelaIndex + 1
                    If CentinelaIndex > NRO_CENTINELA Then Exit Sub
                    
                    ' Si no queda nadie inactivo, me voy
                    CentinelaIndex = GetIdleCentinela(CentinelaIndex)
                    If CentinelaIndex = 0 Then Exit Sub
                    
                End If
            End If
            
        End With
        
    Next LoopC
        
End Sub

Private Function GetIdleCentinela(ByVal StartCheckIndex As Integer) As Integer
'*************************************************
'Author: ZaMa
'Last modified: 07/10/2010
'Returns the index of the first idle centinela found, starting from a given index.
'*************************************************
    Dim index As Long
    
    For index = StartCheckIndex To NRO_CENTINELA
        
        If Not Centinela(index).Activo Then
            GetIdleCentinela = index
            Exit Function
        End If
        
    Next index

End Function

Private Sub CentinelaFinalCheck(ByVal CentiIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Al finalizar el tiempo, se retira y realiza la acción pertinente dependiendo del caso
'03/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************

On Error GoTo Error_Handler

    Dim UserIndex As Integer
    Dim UserName As String
    
    With Centinela(CentiIndex)
    
        UserIndex = .RevisandoUserIndex
    
        If Not UserList(UserIndex).flags.CentinelaOK Then
        
            UserName = UserList(UserIndex).Name
        
            'Logueamos el evento
            Call LogCentinela("Centinela ejecuto y echó a " & UserName & " por uso de macro inasistido.")
            
            'Avisamos a los admins
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> El centinela ha ejecutado a " & UserName & " y lo echó del juego.", FontTypeNames.FONTTYPE_SERVER))
            
            ' Evitamos loguear el logout
            .RevisandoUserIndex = 0
            
            Call WriteShowMessageBox(UserIndex, "Has sido ejecutado por macro inasistido y echado del juego.")
            Call UserDie(UserIndex)
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
        End If
        
        .clave = 0
        .TiempoRestante = 0
        .RevisandoUserIndex = 0
        .Activo = False
        
        If .NpcIndex <> 0 Then
            Call QuitarNPC(.NpcIndex)
            .NpcIndex = 0
        End If
        
    End With
    
    Exit Sub

Error_Handler:

    With Centinela(CentiIndex)
        .clave = 0
        .TiempoRestante = 0
        .RevisandoUserIndex = 0
        .Activo = False
        
        If .NpcIndex Then
            Call QuitarNPC(.NpcIndex)
            .NpcIndex = 0
        End If
    End With
    
    Call LogError("Error en el checkeo del centinela: " & Err.description)
End Sub

Public Sub CentinelaCheckClave(ByVal UserIndex As Integer, ByVal clave As Integer)
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Corrobora la clave que le envia el usuario
'02/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'08/10/2010: ZaMa - Agrego algunos logueos mas coherentes.
'*************************************************

    Dim CentinelaIndex As Byte

    CentinelaIndex = UserList(UserIndex).flags.CentinelaIndex
    
    ' No esta siendo revisado por ningun centinela? Clickeo a alguno?
    If CentinelaIndex = 0 Then
        
        ' Si no clickeo a ninguno, simplemente logueo el evento (Sino hago hablar al centi)
        CentinelaIndex = EsCentinela(UserList(UserIndex).flags.TargetNPC)
        If CentinelaIndex = 0 Then
            Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondió aunque no se le hablaba a él..")
            Exit Sub
        End If
    
    End If
    
    With Centinela(CentinelaIndex)
        If clave = .clave And UserIndex = .RevisandoUserIndex Then
        
            If Not UserList(UserIndex).flags.CentinelaOK Then
        
                UserList(UserIndex).flags.CentinelaOK = True
                Call WriteChatOverHead(UserIndex, "¡Muchas gracias " & UserList(UserIndex).Name & "! Espero no haber sido una molestia.", Npclist(.NpcIndex).Char.CharIndex, vbWhite)
                
                .Activo = False
                Call FlushBuffer(UserIndex)
                
            Else
                'Logueamos el evento
                Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondió más de una vez la contraseña correcta.")
            End If
            
        Else
            
            'Logueamos el evento
            If UserIndex <> .RevisandoUserIndex Then
                Call WriteChatOverHead(UserIndex, "No es a ti a quien estoy hablando, ¿No ves?", Npclist(.NpcIndex).Char.CharIndex, vbWhite)
                Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondió aunque no se le hablaba a él.")
            Else
            
                If Not UserList(UserIndex).flags.CentinelaOK Then
                    ' Clave incorrecta, la reenvio
                    Call CentinelaSendClave(UserIndex, CentinelaIndex)
                    Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondió una clave incorrecta: " & clave & " - Se esperaba : " & .clave)
                Else
                    Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondió una clave incorrecta después de haber respondido una clave correcta.")
                End If
            End If
        End If
    End With
    
End Sub

Public Sub ResetCentinelaInfo()
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Cada determinada cantidad de tiempo, volvemos a revisar
'07/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        
        UserList(LoopC).flags.CentinelaOK = False
        UserList(LoopC).flags.CentinelaIndex = 0
        
    Next LoopC
    
End Sub

Public Sub CentinelaSendClave(ByVal UserIndex As Integer, ByVal CentinelaIndex As Byte)
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Enviamos al usuario la clave vía el personaje centinela
'02/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************

    With Centinela(CentinelaIndex)

        If .NpcIndex = 0 Then Exit Sub
        
        If .RevisandoUserIndex = UserIndex Then
        
            If Not UserList(UserIndex).flags.CentinelaOK Then
                Call WriteChatOverHead(UserIndex, "¡La clave que te he dicho es /CENTINELA " & .clave & ", escríbelo rápido!", Npclist(.NpcIndex).Char.CharIndex, vbGreen)
                Call WriteConsoleMsg(UserIndex, "El centinela intenta llamar tu atención. ¡Respondele rápido!", FontTypeNames.FONTTYPE_CENTINELA)
            Else
                Call WriteChatOverHead(UserIndex, "Te agradezco, pero ya me has respondido. Me retiraré pronto.", CStr(Npclist(.NpcIndex).Char.CharIndex), vbGreen)
            End If
            
        Else
            Call WriteChatOverHead(UserIndex, "No es a ti a quien estoy hablando, ¿No ves?", Npclist(.NpcIndex).Char.CharIndex, vbWhite)
        End If
        
    End With
    
End Sub

Public Sub PasarMinutoCentinela()
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Control del timer. Llamado cada un minuto.
'03/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************
On Error GoTo ErrHandler

    Dim index As Long
    Dim UserIndex As Integer
    Dim IdleCount As Integer
    
    If Not centinelaActivado Then Exit Sub
    
    ' Primero reviso los que estan chequeando usuarios
    For index = 1 To NRO_CENTINELA
    
        With Centinela(index)
            ' Esta activo?
            If .Activo Then
                .TiempoRestante = .TiempoRestante - 1
                
                ' Temrino el tiempo de chequeo?
                If .TiempoRestante = 0 Then
                    Call CentinelaFinalCheck(index)
                Else
                    
                    UserIndex = .RevisandoUserIndex
                
                    'RECORDamos al user que debe escribir
                    If Matematicas.Distancia(Npclist(.NpcIndex).Pos, UserList(UserIndex).Pos) > 5 Then
                        Call WarpCentinela(UserIndex, index)
                    End If
                    
                    'El centinela habla y se manda a consola para que no quepan dudas
                    Call WriteChatOverHead(UserIndex, "¡" & UserList(UserIndex).Name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & .clave & ".", CStr(Npclist(.NpcIndex).Char.CharIndex), vbRed)
                    Call WriteConsoleMsg(UserIndex, "¡" & UserList(UserIndex).Name & ", tienes un minuto más para responder!", FontTypeNames.FONTTYPE_CENTINELA)
                    Call FlushBuffer(UserIndex)
                End If
            Else
            
                ' Lo reseteo aca, para que pueda hablarle al usuario chequeado aunque haya respondido bien.
                If .NpcIndex <> 0 Then
                    If .RevisandoUserIndex <> 0 Then
                        UserList(.RevisandoUserIndex).flags.CentinelaIndex = 0
                        .RevisandoUserIndex = 0
                    End If
                    Call QuitarNPC(.NpcIndex)
                    .NpcIndex = 0
                End If
                
                IdleCount = IdleCount + 1
            End If
            
        End With
    Next index
    
    'Verificamos si ya debemos resetear la lista
    Dim TActual As Long
    TActual = GetTickCount() And &H7FFFFFFF
    
    If checkInterval(centinelaStartTime, TActual, centinelaInterval) Then
        DetenerAsignacion = True ' Espero a que terminen de controlar todos los centinelas
    End If

    ' Si hay algun centinela libre, se fija si no hay trabajadores disponibles para chequear
    If IdleCount <> 0 Then
    
        ' Si es tiempo de resetear flags, chequeo que no quede nadie activo
        If DetenerAsignacion Then
            
            ' No se completaron los ultimos chequeos
            If IdleCount < NRO_CENTINELA Then Exit Sub
            
            ' Resetea todos los flags
            Call ResetCentinelaInfo
            DetenerAsignacion = False
            
            ' Renuevo el contador de reseteo
            RenovarResetTimer
            
        End If
        
        Call GoToNextWorkingChar
        
    End If
    
    Exit Sub
ErrHandler:
    Call LogError("Error en PasarMinutoCentinela. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub WarpCentinela(ByVal UserIndex As Integer, ByVal CentinelaIndex As Byte)
'*************************************************
'Author: Unknown
'Last modified: 02/10/2010
'Inciamos la revisión del usuario UserIndex
'02/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'*************************************************

    With Centinela(CentinelaIndex)

        'Evitamos conflictos de índices
        If .NpcIndex <> 0 Then
            Call QuitarNPC(.NpcIndex)
            .NpcIndex = 0
        End If
        
        ' Spawn it
        .NpcIndex = SpawnNpc(NPC_CENTINELA, UserList(UserIndex).Pos, True, False)
        
        'Si no pudimos crear el NPC, seguimos esperando a poder hacerlo
        If .NpcIndex = 0 Then
            .RevisandoUserIndex = 0
            .Activo = False
        End If
        
    End With
    
End Sub

Public Sub CentinelaUserLogout(ByVal CentinelaIndex As Byte)
'*************************************************
'Author: Unknown
'Last modified: 02/11/2010
'El usuario al que revisabamos se desconectó
'02/10/2010: ZaMa - Adaptado para que funcione mas de un centinela en paralelo.
'02/11/2010: ZaMa - Ahora no loguea que el usuario cerro si puso bien la clave.
'*************************************************
    
    With Centinela(CentinelaIndex)
    
        If .RevisandoUserIndex <> 0 Then
        
            'Logueamos el evento
            If Not UserList(.RevisandoUserIndex).flags.CentinelaOK Then _
                Call LogCentinela("El usuario " & UserList(.RevisandoUserIndex).Name & " se desolgueó al pedirsele la contraseña.")
            
            'Reseteamos y esperamos a otro PasarMinuto para ir al siguiente user
            .clave = 0
            .TiempoRestante = 0
            .RevisandoUserIndex = 0
            .Activo = False
            
            If .NpcIndex <> 0 Then
                Call QuitarNPC(.NpcIndex)
                .NpcIndex = 0
            End If
            
        End If
        
    End With
    
End Sub

Private Sub LogCentinela(ByVal texto As String)
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 03/15/2006
'Loguea un evento del centinela
'*************************************************
On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Centinela.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
Exit Sub

ErrHandler:
End Sub

Public Sub ResetCentinelas()
'*************************************************
'Author: ZaMa
'Last modified: 02/10/2010
'Resetea todos los centinelas
'*************************************************
    Dim index As Long
    Dim UserIndex As Integer
    
    For index = LBound(Centinela) To UBound(Centinela)
        
        With Centinela(index)
            
            ' Si esta activo, reseteo toda la info y quito el npc
            If .Activo Then
                
                .Activo = False
                
                UserIndex = .RevisandoUserIndex
                If UserIndex <> 0 Then
                    UserList(UserIndex).flags.CentinelaIndex = 0
                    UserList(UserIndex).flags.CentinelaOK = False
                    .RevisandoUserIndex = 0
                End If
                
                
                .clave = 0
                .TiempoRestante = 0
                
                If .NpcIndex <> 0 Then
                    Call QuitarNPC(.NpcIndex)
                    .NpcIndex = 0
                End If
                
            End If
            
        End With
    
    Next index
    
    DetenerAsignacion = False
    RenovarResetTimer
    
End Sub

Public Function EsCentinela(ByVal NpcIndex As Integer) As Integer
'*************************************************
'Author: ZaMa
'Last modified: 07/10/2010
'Devuelve True si el indice pertenece a un centinela.
'*************************************************

    Dim index As Long
    
    If NpcIndex = 0 Then Exit Function
    
    For index = 1 To NRO_CENTINELA
    
        If Centinela(index).NpcIndex = NpcIndex Then
            EsCentinela = index
            Exit Function
        End If
        
    Next index

End Function

Private Sub RenovarResetTimer()
'*************************************************
'Author: ZaMa
'Last modified: 07/10/2010
'Renueva el timer que resetea el flag "CentinelaOk" de todos los usuarios.
'*************************************************

    Dim TActual As Long
    TActual = GetTickCount() And &H7FFFFFFF
    
    centinelaInterval = (RandomNumber(0, TIEMPO_PASAR_RANDOM) + TIEMPO_PASAR_BASE) * 60 * 1000
End Sub
