Attribute VB_Name = "wskapiAO"
Option Explicit

''
' Modulo para manejar Winsock
'

#If UsarQueSocket = 1 Then


'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_RECEIVE As Long = &H0
Private Const SD_SEND As Long = &H1
Private Const SD_BOTH As Long = &H2


Private Const MAX_TIEMPOIDLE_COLALLENA = 1 'minutos
Private Const MAX_COLASALIDA_COUNT = 800

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

' ====================================================================================
' ====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

' ====================================================================================
' ====================================================================================

Public SockListen As Long

#End If

' ====================================================================================
' ====================================================================================


Public Sub IniciaWsApi(ByVal hwndParent As Long)
#If UsarQueSocket = 1 Then

Call LogApiSock("IniciaWsApi")
Debug.Print "IniciaWsApi"

#If WSAPI_CREAR_LABEL Then
hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim Desc As String
Call StartWinsock(Desc)

#End If
End Sub

Public Sub LimpiaWsApi(ByVal hWnd As Long)
#If UsarQueSocket = 1 Then

Call LogApiSock("LimpiaWsApi")

If WSAStartedUp Then
    Call EndWinsock
End If

If OldWProc <> 0 Then
    SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
    OldWProc = 0
End If

#If WSAPI_CREAR_LABEL Then
If hWndMsg <> 0 Then
    DestroyWindow hWndMsg
End If
#End If

#End If
End Sub

Public Function BuscaSlotSock(ByVal s As Long, Optional ByVal CacheInd As Boolean = False) As Long
#If UsarQueSocket = 1 Then

On Error GoTo hayerror

BuscaSlotSock = WSAPISock2Usr.Item(CStr(s))
Exit Function

hayerror:
BuscaSlotSock = -1


'
'Dim Pri As Long, Ult As Long, Med As Long
'
'If WSAPISockChacheCant > 0 Then
'    'Busqueda Dicotomica :D
'    Pri = 1
'    Ult = WSAPISockChacheCant
'    Med = Int((Pri + Ult) / 2)
'
'    Do While (Pri <= Ult) And (WSAPISockChache(Med).Sock <> s)
'        If s < WSAPISockChache(Med).Sock Then
'            Ult = Med - 1
'        Else
'            Pri = Med + 1
'        End If
'        Med = Int((Pri + Ult) / 2)
'    Loop
'
'    If Pri <= Ult Then
'        If CacheInd Then
'            BuscaSlotSock = Med
'        Else
'            BuscaSlotSock = WSAPISockChache(Med).Slot
'        End If
'    Else
'        BuscaSlotSock = -1
'    End If
'Else
'    BuscaSlotSock = -1
'End If

#End If
End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
Debug.Print "AgregaSockSlot"
#If (UsarQueSocket = 1) Then

'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("AgregaSlotSock:: sock=" & Sock & " slot=" & Slot)

If WSAPISock2Usr.Count > MaxUsers Then
    'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("Imposible agregarSlotSock (wsapi2usr.count>maxusers)")
    Call CloseSocket(Slot)
    Exit Sub
End If

WSAPISock2Usr.Add CStr(Slot), CStr(Sock)

'Dim Pri As Long, Ult As Long, Med As Long
'Dim LoopC As Long
'
'If WSAPISockChacheCant > 0 Then
'    Pri = 1
'    Ult = WSAPISockChacheCant
'    Med = Int((Pri + Ult) / 2)
'
'    Do While (Pri <= Ult) And (Ult > 1)
'        If Sock < WSAPISockChache(Med).Sock Then
'            Ult = Med - 1
'        Else
'            Pri = Med + 1
'        End If
'        Med = Int((Pri + Ult) / 2)
'    Loop
'
'    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
'    Ult = WSAPISockChacheCant
'    For LoopC = Ult To Pri Step -1
'        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
'    Next LoopC
'    Med = Pri
'Else
'    Med = 1
'End If
'WSAPISockChache(Med).Slot = Slot
'WSAPISockChache(Med).Sock = Sock
'WSAPISockChacheCant = WSAPISockChacheCant + 1

#End If
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)
#If (UsarQueSocket = 1) Then
Dim cant As Long

cant = WSAPISock2Usr.Count
On Error Resume Next
WSAPISock2Usr.Remove CStr(Sock)

Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count

#End If
End Sub



Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If UsarQueSocket = 1 Then

On Error Resume Next

Dim Ret As Long
Dim Tmp As String

Dim s As Long, E As Long
Dim N As Integer
    
Dim Dale As Boolean
Dim UltError As Long


WndProc = 0


If CamaraLenta = 1 Then
    Sleep 1
End If


Select Case msg
Case 1025

    s = wParam
    E = WSAGetSelectEvent(lParam)
    'Debug.Print "Msg: " & msg & " W: " & wParam & " L: " & lParam
    Call LogApiSock("Msg: " & msg & " W: " & wParam & " L: " & lParam)
    
    Select Case E
    Case FD_ACCEPT
            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("FD_ACCEPT")
        If s = SockListen Then
            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("sockLIsten = " & s & ". Llamo a Eventosocketaccept")
            Call EventoSockAccept(s)
        End If
        
'    Case FD_WRITE
'        N = BuscaSlotSock(s)
'        If N < 0 And s <> SockListen Then
'            'Call apiclosesocket(s)
'            call WSApiCloseSocket(s)
'            Exit Function
'        End If
'
'        UserList(N).SockPuedoEnviar = True

'        Call IntentarEnviarDatosEncolados(N)
'
''        Dale = UserList(N).ColaSalida.Count > 0
''        Do While Dale
''            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
''            If Ret <> 0 Then
''                If Ret = WSAEWOULDBLOCK Then
''                    Dale = False
''                Else
''                    'y aca que hacemo' ?? help! i need somebody, help!
''                    Dale = False
''                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
''                End If
''            Else
''            '    Debug.Print "Dato de la cola enviado"
''                UserList(N).ColaSalida.Remove 1
''                Dale = (UserList(N).ColaSalida.Count > 0)
''            End If
''        Loop

    Case FD_READ
        
        N = BuscaSlotSock(s)
        If N < 0 And s <> SockListen Then
            'Call apiclosesocket(s)
            Call WSApiCloseSocket(s)
            Exit Function
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (0))
        
        '4k de buffer
        'buffer externo
        Tmp = Space$(SIZE_RCVBUF)   'si cambias este valor, tambien hacelo mas abajo
                            'donde dice ret = 8192 :)
        
        Ret = recv(s, Tmp, Len(Tmp), 0)
        ' Comparo por = 0 ya que esto es cuando se cierra
        ' "gracefully". (mas abajo)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEMSGSIZE Then
                Debug.Print "WSAEMSGSIZE"
                Ret = SIZE_RCVBUF
            Else
                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                Call LogApiSock("Error en Recv: N=" & N & " S=" & s & " Str=" & GetWSAErrorString(UltError))
                
                'no hay q llamar a CloseSocket() directamente,
                'ya q pueden abusar de algun error para
                'desconectarse sin los 10segs. CREEME.
            '    Call C l o s e Socket(N)
            
                Call CloseSocketSL(N)
                Call Cerrar_Usuario(N)
                Exit Function
            End If
        ElseIf Ret = 0 Then
            Call CloseSocketSL(N)
            Call Cerrar_Usuario(N)
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT))
        
        Tmp = Left(Tmp, Ret)
        
        'Call LogApiSock("WndProc:FD_READ:N=" & N & ":TMP=" & Tmp)
        
        Call EventoSockRead(N, Tmp)
        
    Case FD_CLOSE
        N = BuscaSlotSock(s)
        If s <> SockListen Then Call apiclosesocket(s)
        
        Call LogApiSock("WndProc:FD_CLOSE:N=" & N & ":Err=" & WSAGetAsyncError(lParam))
        
        If N > 0 Then
            Call BorraSlotSock(UserList(N).ConnID)
            UserList(N).ConnID = -1
            UserList(N).ConnIDValida = False
            Call EventoSockClose(N)
        End If
        
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
End Select

#End If
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByVal str As String, Optional Encolar As Boolean = True) As Long
#If UsarQueSocket = 1 Then

'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("WsApiEnviar:: slot=" & Slot & " str=" & str & " len(str)=" & Len(str) & " encolar=" & Encolar)

Dim Ret As String
Dim UltError As Long
Dim Retorno As Long

Retorno = 0

'Debug.Print ">>>> " & str

If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
    If ((UserList(Slot).ColaSalida.Count = 0)) Or (Not Encolar) Then
        Ret = send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEWOULDBLOCK Then
                UserList(Slot).SockPuedoEnviar = False
                If Encolar Then
                    UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
                    'LogCustom ("Encolados datos:" & str)
                End If
            End If
            Retorno = UltError
        End If
    Else
        If UserList(Slot).ColaSalida.Count < MAX_COLASALIDA_COUNT Or UserList(Slot).Counters.IdleCount < MAX_TIEMPOIDLE_COLALLENA Then
            UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
            
        Else
            Retorno = -1
        End If
    End If
ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
    If Not UserList(Slot).Counters.Saliendo Then
        Retorno = -1
    End If
End If

WsApiEnviar = Retorno

#End If
End Function


Public Sub LogCustom(ByVal str As String)
#If (UsarQueSocket = 1) Then

On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\custom.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & "(" & Timer & ") " & str
Close #nfile

Exit Sub

errhandler:

#End If
End Sub


Public Sub LogApiSock(ByVal str As String)
#If (UsarQueSocket = 1) Then

On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

#End If
End Sub


Public Sub IntentarEnviarDatosEncolados(ByVal N As Integer)
#If UsarQueSocket = 1 Then

Dim Dale As Boolean
Dim Ret As Long

Dale = UserList(N).ColaSalida.Count > 0
Do While Dale
    Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
    If Ret <> 0 Then
        If Ret = WSAEWOULDBLOCK Then
            Dale = False
        Else
            'y aca que hacemo' ?? help! i need somebody, help!
            Dale = False
            Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
            Call LogApiSock("IntentarEnviarDatosEncolados: N=" & N & " " & GetWSAErrorString(Ret))
            Call CloseSocketSL(N)
            Call Cerrar_Usuario(N)
        End If
    Else
    '    Debug.Print "Dato de la cola enviado"
        UserList(N).ColaSalida.Remove 1
        Dale = (UserList(N).ColaSalida.Count > 0)
    End If
Loop

#End If
End Sub


Public Sub EventoSockAccept(ByVal SockID As Long)
#If UsarQueSocket = 1 Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================
    
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim tStr As String
    
    Tam = sockaddr_size
    
    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
'Modificado por Maraxus
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
    Ret = accept(SockID, sa, Tam)

    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If

    'If Ret = INVALID_SOCKET Then
    '    If Err.LastDllError = 11002 Then
    '        ' We couldn't decide if to accept or reject the connection
    '        'Force reject so we can get it out of the queue
    '        LogCustom ("Pre WSAAccept CallbackData=1")
    '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
    '        LogCustom ("WSAccept Callbackdata 1, devuelve " & Ret)
    '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
    '    Else
    '        i = Err.LastDllError
    '        LogCustom ("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    '        Exit Sub
    '    End If
    'End If

    NuevoSock = Ret
    
    'Seteamos el tamaño del buffer de entrada a 512 bytes
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    'Seteamos el tamaño del buffer de salida a 1 Kb
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If

    If False Then
    'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
        tStr = "ERRLimite de conexiones para su IP alcanzado." & ENDC
        Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser ' Nuevo indice
    
    If NewIndex <= MaxUsers Then
        
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                tStr = "ERRSu IP se encuentra bloqueada en este servidor." & ENDC
                Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
                'Call SecurityIp.IpRestarConexion(sa.sin_addr)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        UserList(NewIndex).SockPuedoEnviar = True
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        Set UserList(NewIndex).ColaSalida = New Collection
        
        Call AgregaSlotSock(NuevoSock, NewIndex)
    Else
        tStr = "ERRServer lleno." & ENDC
        Dim AAA As Long
        AAA = send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        'Call SecurityIp.IpRestarConexion(sa.sin_addr)
        Call WSApiCloseSocket(NuevoSock)
    End If
    
#End If
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos As String)
#If UsarQueSocket = 1 Then

Dim t() As String
Dim LoopC As Long

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

t = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(Slot).RDBuffer = t(UBound(t))
    
    For LoopC = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(LoopC) <> "" Then If Not UserList(Slot).CommandsBuffer.Push(t(LoopC)) Then Call CloseSocket(Slot)
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If

#End If
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
#If UsarQueSocket = 1 Then

    'Es el mismo user al que está revisando el centinela??
    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
    If Centinela.RevisandoUserIndex = Slot Then _
        Call modCentinela.CentinelaUserLogout
    
    If UserList(Slot).flags.UserLogged Then
        Call CloseSocketSL(Slot)
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
#End If
End Sub


Public Sub WSApiReiniciarSockets()
#If UsarQueSocket = 1 Then
Dim i As Long
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If
        
        'Call ResetUserSlot(i)
    Next i
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)
    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
    Next i
    
    LastUser = 1
    NumUsers = 0
    
    Call LimpiaWsApi(frmMain.hWnd)
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")


#End If
End Sub


Public Sub WSApiCloseSocket(ByVal Socket As Long)
#If UsarQueSocket = 1 Then
Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
Call ShutDown(Socket, SD_BOTH)
#End If
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
#If UsarQueSocket = 1 Then
    Dim sa As sockaddr
    
    'Check if we were requested to force reject

    If dwCallbackData = 1 Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If
    
     'Get the address

    CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen

    
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If

    CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
#End If
End Function
