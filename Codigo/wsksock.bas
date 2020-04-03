Attribute VB_Name = "WSKSOCK"
'**************************************************************
' WSKSOCK.bas
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

'date stamp: sept 1, 1996 (for version control, please don't remove)

'Visual Basic 4.0 Winsock "Header"
'   Alot of the information contained inside this file was originally
'   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
'   modified in some way.
'
'Disclaimer: This file is public domain, updated periodically by
'   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
'   Neither myself(Topaz) or anyone related to alt.programming.winsock
'   may be held liable for its use, or misuse.
'
'Declare check Aug 27, 1996. (Topaz, SigSegV@mail.utexas.edu)
'   All 16 bit declarations appear correct, even the odd ones that
'   pass longs inplace of in_addr and char buffers. 32 bit functions
'   also appear correct. Some are declared to return integers instead of
'   longs (breaking MS's rules.) however after testing these functions I
'   have come to the conclusion that they do not work properly when declared
'   following MS's rules.
'
'NOTES:
'   (1) I have never used WS_SELECT (select), therefore I must warn that I do
'       not know if fd_set and timeval are properly defined.
'   (2) Alot of the functions are declared with "buf as any", when calling these
'       functions you may either pass strings, byte arrays or UDT's. For 32bit I
'       I recommend Byte arrays and the use of memcopy to copy the data back out
'   (3) The async functions (wsaAsync*) require the use of a message hook or
'       message window control to capture messages sent by the winsock stack. This
'       is not to be confused with a CallBack control, The only function that uses
'       callbacks is WSASetBlockingHook()
'   (4) Alot of "helper" functions are provided in the file for various things
'       before attempting to figure out how to call a function, look and see if
'       there is already a helper function for it.
'   (5) Data types (hostent etc) have kept there 16bit definitions, even under 32bit
'       windows due to the problem of them not working when redfined following the
'       suggested rules.
Option Explicit

Public Const FD_SETSIZE = 64

Type fd_set

    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer

End Type

Type timeval

    tv_sec As Long
    tv_usec As Long

End Type

Type HostEnt

    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long

End Type

Public Const hostent_size = 16

Type servent

    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long

End Type

Public Const servent_size = 14

Type protoent

    p_name As Long
    p_aliases As Long
    p_proto As Integer

End Type

Public Const protoent_size = 10

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17
Public Const INADDR_NONE = &HFFFFFFFF
Public Const INADDR_ANY = &H0

Type sockaddr

    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8

End Type

Public Const sockaddr_size = 16

Public saZero As sockaddr

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType

    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long

End Type

'Agregado por Maraxus
Type WSABUF

    dwBufferLen As Long
    lpBuffer    As Long

End Type

'Agregado por Maraxus
Type FLOWSPEC

    TokenRate           As Long     'In Bytes/sec
    TokenBucketSize     As Long     'In Bytes
    PeakBandwidth       As Long     'In Bytes/sec
    Latency             As Long     'In microseconds
    DelayVariation      As Long     'In microseconds
    ServiceType         As Integer  'Guaranteed, Predictive,
    'Best Effort, etc.
    MaxSduSize          As Long     'In Bytes
    MinimumPolicedSize  As Long     'In Bytes

End Type

'Agregado por Maraxus
Public Const WSA_FLAG_OVERLAPPED = &H1

'Agregados por Maraxus
Public Const CF_ACCEPT = &H0
Public Const CF_REJECT = &H1

'Agregado por Maraxus
Public Const SD_RECEIVE As Long = &H0&
Public Const SD_SEND    As Long = &H1&
Public Const SD_BOTH    As Long = &H2&
Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const MAXGETHOSTSTRUCT = 1024
Public Const AF_INET = 2
Public Const PF_INET = 2

Type LingerType

    l_onoff As Integer
    l_linger As Integer

End Type

' Windows Sockets definitions of regular Microsoft C error constants
Public Const WSAEINTR = 10004
Public Const WSAEBADF = 10009
Public Const WSAEACCES = 10013
Public Const WSAEFAULT = 10014
Public Const WSAEINVAL = 10022
Public Const WSAEMFILE = 10024

' Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK = 10035
Public Const WSAEINPROGRESS = 10036
Public Const WSAEALREADY = 10037
Public Const WSAENOTSOCK = 10038
Public Const WSAEDESTADDRREQ = 10039
Public Const WSAEMSGSIZE = 10040
Public Const WSAEPROTOTYPE = 10041
Public Const WSAENOPROTOOPT = 10042
Public Const WSAEPROTONOSUPPORT = 10043
Public Const WSAESOCKTNOSUPPORT = 10044
Public Const WSAEOPNOTSUPP = 10045
Public Const WSAEPFNOSUPPORT = 10046
Public Const WSAEAFNOSUPPORT = 10047
Public Const WSAEADDRINUSE = 10048
Public Const WSAEADDRNOTAVAIL = 10049
Public Const WSAENETDOWN = 10050
Public Const WSAENETUNREACH = 10051
Public Const WSAENETRESET = 10052
Public Const WSAECONNABORTED = 10053
Public Const WSAECONNRESET = 10054
Public Const WSAENOBUFS = 10055
Public Const WSAEISCONN = 10056
Public Const WSAENOTCONN = 10057
Public Const WSAESHUTDOWN = 10058
Public Const WSAETOOMANYREFS = 10059
Public Const WSAETIMEDOUT = 10060
Public Const WSAECONNREFUSED = 10061
Public Const WSAELOOP = 10062
Public Const WSAENAMETOOLONG = 10063
Public Const WSAEHOSTDOWN = 10064
Public Const WSAEHOSTUNREACH = 10065
Public Const WSAENOTEMPTY = 10066
Public Const WSAEPROCLIM = 10067
Public Const WSAEUSERS = 10068
Public Const WSAEDQUOT = 10069
Public Const WSAESTALE = 10070
Public Const WSAEREMOTE = 10071

' Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY = 10091
Public Const WSAVERNOTSUPPORTED = 10092
Public Const WSANOTINITIALISED = 10093
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004
Public Const WSANO_ADDRESS = 11004

'---ioctl Constants
Public Const FIONREAD = &H8004667F

Public Const FIONBIO = &H8004667E

Public Const FIOASYNC = &H8004667D

'---Windows System Functions
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'---async notification constants
Public Const SOL_SOCKET = &HFFFF&
Public Const SO_LINGER = &H80&
Public Const SO_RCVBUFFER = &H1002&             ' Agregado por Maraxus
Public Const SO_SNDBUFFER = &H1001&              ' Agregado por Maraxus
Public Const SO_CONDITIONAL_ACCEPT = &H3002&    ' Agregado por Maraxus
Public Const FD_READ = &H1&
Public Const FD_WRITE = &H2&
Public Const FD_OOB = &H4&
Public Const FD_ACCEPT = &H8&
Public Const FD_CONNECT = &H10&
Public Const FD_CLOSE = &H20&

'---SOCKET FUNCTIONS
Public Declare Function accept Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, AddrLen As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function apiclosesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal S As Long) As Long
Public Declare Function connect Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal S As Long, ByVal Cmd As Long, argp As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long

Public Declare Function listen Lib "ws2_32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
Public Declare Function recv Lib "ws2_32.dll" (ByVal S As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
Public Declare Function ws_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
Public Declare Function send Lib "ws2_32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function ShutDown Lib "ws2_32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long

'---DATABASE FUNCTIONS
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Long, ByVal proto As String) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long

'---WINDOWS EXTENSIONS
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Sub WSASetLastError Lib "ws2_32.dll" (ByVal iError As Long)
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAIsBlocking Lib "ws2_32.dll" () As Long
Public Declare Function WSAUnhookBlockingHook Lib "ws2_32.dll" () As Long
Public Declare Function WSASetBlockingHook Lib "ws2_32.dll" (ByVal lpBlockFunc As Long) As Long
Public Declare Function WSACancelBlockingCall Lib "ws2_32.dll" () As Long
Public Declare Function WSAAsyncGetServByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetServByPort Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByNumber Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSACancelAsyncRequest Lib "ws2_32.dll" (ByVal hAsyncTaskHandle As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSARecvEx Lib "ws2_32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long

'Agregado por Maraxus
Declare Function WSAAccept Lib "ws2_32.dll" (ByVal S As Long, pSockAddr As sockaddr, AddrLen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long

Public Const SOMAXCONN As Long = &H7FFFFFFF            ' Agregado por Maraxus

'SOME STUFF I ADDED
Public MySocket%
Public SockReadBuffer$

Public Const WSA_NoName = "Unknown"

Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled

Public Function StartWinsock(sDescription As String) As Boolean

    Dim StartupData As WSADataType

    If Not WSAStartedUp Then

        If Not WSAStartup(&H202, StartupData) Then  'Use sockets v2.2 instead of 1.1 (Maraxus)
            WSAStartedUp = True
            'Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
            'Debug.Print "If wVersion == 257 then everything is kewl"
            'Debug.Print "szDescription="; StartupData.szDescription
            'Debug.Print "szSystemStatus="; StartupData.szSystemStatus
            'Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
        
        Else
            WSAStartedUp = False

        End If

    End If

    StartWinsock = WSAStartedUp

End Function

Public Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long

    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetAsyncBufLen = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetAsyncBufLen = lParam And &HFFFF&
    End If

End Function

Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer

    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If

End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Public Function AddrToIP(ByVal AddrOrIP$) As String

    Dim T() As String
    Dim Tmp As String

    Tmp = GetAscIP(GetHostByNameAlias(AddrOrIP$))
    
    T = Split(Tmp, ".")
    
    AddrToIP = T(3) & "." & T(2) & "." & T(1) & "." & T(0)

End Function

Function ConnectSock(ByVal Host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long

    Dim S&, SelectOps&, dummy&
    Dim sockin As sockaddr

    SockReadBuffer$ = vbNullString
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)

    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(Host$)

    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    retIpPort$ = GetAscIP$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    S = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)

    If S < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    If SetSockLinger(S, 1, 0) = SOCKET_ERROR Then
        
        If S > 0 Then
            dummy = apiclosesocket(S)
        End If

        ConnectSock = INVALID_SOCKET
        
        Exit Function

    End If

    If Not Async Then
        
        If Not connect(S, sockin, sockaddr_size) = 0 Then
            If S > 0 Then
                dummy = apiclosesocket(S)
            End If

            ConnectSock = INVALID_SOCKET
            
            Exit Function

        End If

        If HWndToMsg <> 0 Then
            SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

            If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
                
                If S > 0 Then
                    dummy = apiclosesocket(S)
                End If

                ConnectSock = INVALID_SOCKET
                
                Exit Function

            End If

        End If

    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

        If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            
            If S > 0 Then
                dummy = apiclosesocket(S)
            End If

            ConnectSock = INVALID_SOCKET
            
            Exit Function

        End If

        If connect(S, sockin, sockaddr_size) <> -1 Then
            
            If S > 0 Then
                dummy = apiclosesocket(S)
            End If

            ConnectSock = INVALID_SOCKET
            
            Exit Function

        End If

    End If

    ConnectSock = S

End Function

Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long

    Dim Linger As LingerType

    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime

    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    
    Else

        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger
        End If

    End If

End Function

Sub EndWinsock()

    Dim ret&

    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If

    ret = WSACleanup()
    
    WSAStartedUp = False

End Sub

Public Function GetAscIP(ByVal inn As Long) As String

    Dim nStr&
    Dim lpStr&
    Dim retString$

    retString = String(32, 0)
    lpStr = inet_ntoa(inn)

    If lpStr Then
        nStr = lstrlen(lpStr)

        If nStr > 32 Then nStr = 32
        
        Call MemCopy(ByVal retString, ByVal lpStr, nStr)
        
        retString = Left$(retString, nStr)
        
        GetAscIP = retString
    
    Else
        GetAscIP = "255.255.255.255"
    End If

End Function

Public Function GetHostByAddress(ByVal addr As Long) As String

    Dim phe&
    Dim heDestHost As HostEnt
    Dim HostName$

    phe = gethostbyaddr(addr, 4, PF_INET)

    If phe Then
        Call MemCopy(heDestHost, ByVal phe, hostent_size)
        
        HostName = String$(256, 0)
        
        Call MemCopy(ByVal HostName, ByVal heDestHost.h_name, 256)
        
        GetHostByAddress = Left$(HostName, InStr(HostName, Chr$(0)) - 1)
        
    Else
        GetHostByAddress = WSA_NoName
        
    End If

End Function

'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal HostName$) As Long

    'Return IP address as a long, in network byte order
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&

    retIP = inet_addr(HostName$)

    If retIP = INADDR_NONE Then
        phe = gethostbyname(HostName$)

        If phe <> 0 Then
            Call MemCopy(heDestHost, ByVal phe, hostent_size)
            Call MemCopy(addrList, ByVal heDestHost.h_addr_list, 4)
            Call MemCopy(retIP, ByVal addrList, heDestHost.h_length)
        Else
            retIP = INADDR_NONE
        End If

    End If

    GetHostByNameAlias = retIP

End Function

'returns your local machines name
Public Function GetLocalHostName() As String

    Dim sName$
        sName = String(256, 0)

    If gethostname(sName, 256) Then
        sName = WSA_NoName
    
    Else

        If InStr(sName, Chr$(0)) Then
            sName = Left$(sName, InStr(sName, Chr$(0)) - 1)
        End If

    End If

    GetLocalHostName = sName

End Function

Public Function GetPeerAddress(ByVal S&) As String

    Dim AddrLen&
    Dim sa As sockaddr

    AddrLen = sockaddr_size

    If getpeername(S, sa, AddrLen) Then
        GetPeerAddress = vbNullString
    Else
        GetPeerAddress = SockAddressToString(sa)
    End If

End Function

Public Function GetPortFromString(ByVal PortStr$) As Long

    'sometimes users provide ports outside the range of a VB
    'integer, so this function returns an integer for a string
    'just to keep an error from happening, it converts the
    'number to a negative if needed
    If val(PortStr$) > 32767 Then
        GetPortFromString = CInt(val(PortStr$) - &H10000)
    Else
        GetPortFromString = val(PortStr$)
    End If

    If Err Then GetPortFromString = 0

End Function

Function GetProtocolByName(ByVal Protocol$) As Long

    Dim tmpShort&
    Dim ppe&
    Dim peDestProt As protoent

    ppe = getprotobyname(Protocol)

    If ppe Then
        Call MemCopy(peDestProt, ByVal ppe, protoent_size)
        GetProtocolByName = peDestProt.p_proto
    
    Else
        tmpShort = val(Protocol)

        If tmpShort Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = SOCKET_ERROR
        End If

    End If

End Function

Function GetServiceByName(ByVal service$, ByVal Protocol$) As Long

    Dim Serv&
    Dim pse&
    Dim seDestServ As servent

    pse = getservbyname(service, Protocol)

    If pse Then
        Call MemCopy(seDestServ, ByVal pse, servent_size)
        GetServiceByName = seDestServ.s_port
    
    Else
        Serv = val(service)

        If Serv Then
            GetServiceByName = htons(Serv)
        Else
            GetServiceByName = INVALID_SOCKET
        End If

    End If

End Function

'this function DOES work on 16 and 32 bit systems

Function GetSockAddress(ByVal S&) As String

    Dim AddrLen&
    Dim ret&
    Dim sa As sockaddr
    Dim szRet$

    szRet = String(32, 0)
    AddrLen = sockaddr_size

    If getsockname(S, sa, AddrLen) Then
        GetSockAddress = vbNullString
    Else
        GetSockAddress = SockAddressToString(sa)
    End If

End Function

'this function should work on 16 and 32 bit systems
Function GetWSAErrorString(ByVal errnum&) As String

    On Error Resume Next

    Select Case errnum

        Case 10004
            GetWSAErrorString = "Interrupted system call."

        Case 10009
            GetWSAErrorString = "Bad file number."

        Case 10013
            GetWSAErrorString = "Permission Denied."

        Case 10014
            GetWSAErrorString = "Bad Address."

        Case 10022
            GetWSAErrorString = "Invalid Argument."

        Case 10024
            GetWSAErrorString = "Too many open files."

        Case 10035
            GetWSAErrorString = "Operation would block."

        Case 10036
            GetWSAErrorString = "Operation now in progress."

        Case 10037
            GetWSAErrorString = "Operation already in progress."

        Case 10038
            GetWSAErrorString = "Socket operation on nonsocket."

        Case 10039
            GetWSAErrorString = "Destination address required."

        Case 10040
            GetWSAErrorString = "Message too long."

        Case 10041
            GetWSAErrorString = "Protocol wrong type for socket."

        Case 10042
            GetWSAErrorString = "Protocol not available."

        Case 10043
            GetWSAErrorString = "Protocol not supported."

        Case 10044
            GetWSAErrorString = "Socket type not supported."

        Case 10045
            GetWSAErrorString = "Operation not supported on socket."

        Case 10046
            GetWSAErrorString = "Protocol family not supported."

        Case 10047
            GetWSAErrorString = "Address family not supported by protocol family."

        Case 10048
            GetWSAErrorString = "Address already in use."

        Case 10049
            GetWSAErrorString = "Can't assign requested address."

        Case 10050
            GetWSAErrorString = "Network is down."

        Case 10051
            GetWSAErrorString = "Network is unreachable."

        Case 10052
            GetWSAErrorString = "Network dropped connection."

        Case 10053
            GetWSAErrorString = "Software caused connection abort."

        Case 10054
            GetWSAErrorString = "Connection reset by peer."

        Case 10055
            GetWSAErrorString = "No buffer space available."

        Case 10056
            GetWSAErrorString = "Socket is already connected."

        Case 10057
            GetWSAErrorString = "Socket is not connected."

        Case 10058
            GetWSAErrorString = "Can't send after socket shutdown."

        Case 10059
            GetWSAErrorString = "Too many references: can't splice."

        Case 10060
            GetWSAErrorString = "Connection timed out."

        Case 10061
            GetWSAErrorString = "Connection refused."

        Case 10062
            GetWSAErrorString = "Too many levels of symbolic links."

        Case 10063
            GetWSAErrorString = "File name too long."

        Case 10064
            GetWSAErrorString = "Host is down."

        Case 10065
            GetWSAErrorString = "No route to host."

        Case 10066
            GetWSAErrorString = "Directory not empty."

        Case 10067
            GetWSAErrorString = "Too many processes."

        Case 10068
            GetWSAErrorString = "Too many users."

        Case 10069
            GetWSAErrorString = "Disk quota exceeded."

        Case 10070
            GetWSAErrorString = "Stale NFS file handle."

        Case 10071
            GetWSAErrorString = "Too many levels of remote in path."

        Case 10091
            GetWSAErrorString = "Network subsystem is unusable."

        Case 10092
            GetWSAErrorString = "Winsock DLL cannot support this application."

        Case 10093
            GetWSAErrorString = "Winsock not initialized."

        Case 10101
            GetWSAErrorString = "Disconnect."

        Case 11001
            GetWSAErrorString = "Host not found."

        Case 11002
            GetWSAErrorString = "Nonauthoritative host not found."

        Case 11003
            GetWSAErrorString = "Nonrecoverable error."

        Case 11004
            GetWSAErrorString = "Valid name, no data RECORD of requested type."

        Case Else:

    End Select

End Function

'this function DOES work on 16 and 32 bit systems
Function IpToAddr(ByVal AddrOrIP$) As String

    On Error Resume Next

    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))

    If Err Then IpToAddr = WSA_NoName

End Function

'this function DOES work on 16 and 32 bit systems
Function IrcGetAscIp(ByVal IPL$) As String

    'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
    'the kind that would be parsed out of a DCC command string
    On Error GoTo IrcGetAscIPError:

    Dim lpStr&
    Dim nStr&
    Dim retString$
    Dim inn&

    If val(IPL) > 2147483647 Then
        inn = val(IPL) - 4294967296#
    Else
        inn = val(IPL)
    End If

    inn = ntohl(inn)
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)

    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
    End If

    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    
    Call MemCopy(ByVal retString, ByVal lpStr, nStr)
    
    retString = Left$(retString, nStr)
    
    IrcGetAscIp = retString
    
    Exit Function
    
IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"
    Exit Function
    Resume

End Function

Public Function GetLongIp(ByVal IPS As String) As Long
    GetLongIp = inet_addr(IPS)

End Function

Function IrcGetLongIp(ByVal AscIp$) As String
'this function converts an ascii ip string into a long ip in network byte order
'and stick it in a string suitable for use in a DCC command.

On Error GoTo IrcGetLongIpError:

    Dim inn&

    inn = inet_addr(AscIp)
    inn = htonl(inn)

    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
    Else
        IrcGetLongIp = CVar(inn)
        Exit Function
    End If

    Exit Function
    
IrcGetLongIpError:
    IrcGetLongIp = "0"
    Exit Function
    Resume

End Function

'this function should work on 16 and 32 bit systems

Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, ByVal Enlazar As String) As Long

    Dim S&, dummy&

    Dim SelectOps&

    Dim sockin As sockaddr

    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)

    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If

    If LenB(Enlazar) = 0 Then
        sockin.sin_addr = htonl(INADDR_ANY)
    Else
        sockin.sin_addr = inet_addr(Enlazar)
    End If

    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If

    S = Socket(PF_INET, SOCK_STREAM, 0)

    If S < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    'Agregado por Maraxus
    'If setsockopt(S, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, True, 2) Then
    '    LogApiSock ("Error seteando conditional accept")
    '    Debug.Print "Error seteando conditional accept"
    'Else
    '    LogApiSock ("Conditional accept seteado")
    '    Debug.Print "Conditional accept seteado ^^"
    'End If
    
    If bind(S, sockin, sockaddr_size) Then
        
        If S > 0 Then
            dummy = apiclosesocket(S)
        End If

        ListenForConnect = INVALID_SOCKET
        
        Exit Function

    End If

    SelectOps = FD_READ Or FD_CLOSE Or FD_ACCEPT

    If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        
        If S > 0 Then
            dummy = apiclosesocket(S)
        End If

        ListenForConnect = SOCKET_ERROR
        Exit Function

    End If
    
    'If listen(s, 5) Then
    If listen(S, SOMAXCONN) Then
        
        If S > 0 Then
            dummy = apiclosesocket(S)
        End If

        ListenForConnect = INVALID_SOCKET
        
        Exit Function

    End If

    ListenForConnect = S

End Function

Public Function kSendData(ByVal S&, vMessage As Variant) As Long

    Dim TheMsg() As Byte, sTemp$

    TheMsg = vbNullString

    Select Case VarType(vMessage)

        Case 8209   'byte array
            sTemp = vMessage
            TheMsg = sTemp

        Case 8      'string, if we recieve a string, its assumed we are linemode
            sTemp = StrConv(vMessage, vbFromUnicode)

        Case Else
            sTemp = StrConv(CStr(vMessage), vbFromUnicode)

    End Select

    TheMsg = sTemp

    If UBound(TheMsg) > -1 Then
        kSendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If

End Function

Public Function SockAddressToString(sa As sockaddr) As String
    SockAddressToString = GetAscIP(sa.sin_addr) & ":" & ntohs(sa.sin_port)
End Function

Public Function WSAMakeSelectReply(TheEvent%, TheError%) As Long
    WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)
End Function
