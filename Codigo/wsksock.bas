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

#If UsarQueSocket = 1 Then

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

    'Agregados por Maraxus
    Public Const CF_ACCEPT = &H0

    Public Const CF_REJECT = &H1

    'Agregado por Maraxus
    Public Const INVALID_SOCKET = -1

    Public Const SOCKET_ERROR = -1

    Public Const SOCK_STREAM = 1

    Public Const SOCK_DGRAM = 2

    Public Const AF_INET = 2

    Public Const PF_INET = 2

    Type LingerType

        l_onoff As Integer
        l_linger As Integer

    End Type

    ' Windows Sockets definitions of regular Microsoft C error constants
    Global Const WSAEINTR = 10004

    Global Const WSAEBADF = 10009

    Global Const WSAEACCES = 10013

    Global Const WSAEFAULT = 10014

    Global Const WSAEINVAL = 10022

    Global Const WSAEMFILE = 10024

    ' Windows Sockets definitions of regular Berkeley error constants
    Global Const WSAEWOULDBLOCK = 10035

    Global Const WSAEINPROGRESS = 10036

    Global Const WSAEALREADY = 10037

    Global Const WSAENOTSOCK = 10038

    Global Const WSAEDESTADDRREQ = 10039

    Global Const WSAEMSGSIZE = 10040

    Global Const WSAEPROTOTYPE = 10041

    Global Const WSAENOPROTOOPT = 10042

    Global Const WSAEPROTONOSUPPORT = 10043

    Global Const WSAESOCKTNOSUPPORT = 10044

    Global Const WSAEOPNOTSUPP = 10045

    Global Const WSAEPFNOSUPPORT = 10046

    Global Const WSAEAFNOSUPPORT = 10047

    Global Const WSAEADDRINUSE = 10048

    Global Const WSAEADDRNOTAVAIL = 10049

    Global Const WSAENETDOWN = 10050

    Global Const WSAENETUNREACH = 10051

    Global Const WSAENETRESET = 10052

    Global Const WSAECONNABORTED = 10053

    Global Const WSAECONNRESET = 10054

    Global Const WSAENOBUFS = 10055

    Global Const WSAEISCONN = 10056

    Global Const WSAENOTCONN = 10057

    Global Const WSAESHUTDOWN = 10058

    Global Const WSAETOOMANYREFS = 10059

    Global Const WSAETIMEDOUT = 10060

    Global Const WSAECONNREFUSED = 10061

    Global Const WSAELOOP = 10062

    Global Const WSAENAMETOOLONG = 10063

    Global Const WSAEHOSTDOWN = 10064

    Global Const WSAEHOSTUNREACH = 10065

    Global Const WSAENOTEMPTY = 10066

    Global Const WSAEPROCLIM = 10067

    Global Const WSAEUSERS = 10068

    Global Const WSAEDQUOT = 10069

    Global Const WSAESTALE = 10070

    Global Const WSAEREMOTE = 10071

    ' Extended Windows Sockets error constant definitions
    Global Const WSASYSNOTREADY = 10091

    Global Const WSAVERNOTSUPPORTED = 10092

    Global Const WSANOTINITIALISED = 10093

    Global Const WSAHOST_NOT_FOUND = 11001

    Global Const WSATRY_AGAIN = 11002

    Global Const WSANO_RECOVERY = 11003

    Global Const WSANO_DATA = 11004

    Global Const WSANO_ADDRESS = 11004
    
    #If Win16 Then

        '---Windows System functions
        Public Declare Function PostMessage _
                       Lib "User" (ByVal hWnd As Integer, _
                                   ByVal wMsg As Integer, _
                                   ByVal wParam As Integer, _
                                   lParam As Any) As Integer

        Public Declare Sub MemCopy _
                       Lib "Kernel" _
                       Alias "hmemcpy" (Dest As Any, _
                                        Src As Any, _
                                        ByVal cb&)

        Public Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer

        '---async notification constants
        Public Const SOL_SOCKET = &HFFFF

        Public Const SO_LINGER = &H80

        Public Const SO_RCVBUFFER = &H1002              ' Agregado por Maraxus

        Public Const SO_SNDBUFFER = &H1001              ' Agregado por Maraxus

        Public Const FD_READ = &H1

        Public Const FD_WRITE = &H2

        Public Const FD_ACCEPT = &H8

        Public Const FD_CONNECT = &H10

        Public Const FD_CLOSE = &H20

        '---SOCKET FUNCTIONS
        Public Declare Function accept _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         addr As sockaddr, _
                                         AddrLen As Integer) As Integer

        Public Declare Function bind _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         addr As sockaddr, _
                                         ByVal namelen As Integer) As Integer

        Public Declare Function apiclosesocket _
                       Lib "ws2_32.DLL" _
                       Alias "closesocket" (ByVal S As Integer) As Integer

        Public Declare Function connect _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         addr As sockaddr, _
                                         ByVal namelen As Integer) As Integer

        Public Declare Function ioctlsocket _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByVal Cmd As Long, _
                                         argp As Long) As Integer

        Public Declare Function getpeername _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         sName As sockaddr, _
                                         namelen As Integer) As Integer

        Public Declare Function getsockname _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         sName As sockaddr, _
                                         namelen As Integer) As Integer

        Public Declare Function getsockopt _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByVal level As Integer, _
                                         ByVal optname As Integer, _
                                         optval As Any, _
                                         optlen As Integer) As Integer

        Public Declare Function htonl Lib "ws2_32.DLL" (ByVal hostlong As Long) As Long

        Public Declare Function htons _
                       Lib "ws2_32.DLL" (ByVal hostshort As Integer) As Integer

        Public Declare Function inet_addr Lib "ws2_32.DLL" (ByVal cp As String) As Long

        Public Declare Function inet_ntoa Lib "ws2_32.DLL" (ByVal inn As Long) As Long

        Public Declare Function listen _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByVal backlog As Integer) As Integer

        Public Declare Function ntohl Lib "ws2_32.DLL" (ByVal netlong As Long) As Long

        Public Declare Function ntohs _
                       Lib "ws2_32.DLL" (ByVal netshort As Integer) As Integer

        Public Declare Function recv _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByRef buf As Any, _
                                         ByVal buflen As Integer, _
                                         ByVal flags As Integer) As Integer

        Public Declare Function recvfrom _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer, _
                                         ByVal flags As Integer, _
                                         from As sockaddr, _
                                         fromlen As Integer) As Integer

        Public Declare Function ws_select _
                       Lib "ws2_32.DLL" _
                       Alias "select" (ByVal nfds As Integer, _
                                       readfds As Any, _
                                       writefds As Any, _
                                       exceptfds As Any, _
                                       timeout As timeval) As Integer

        Public Declare Function send _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer, _
                                         ByVal flags As Integer) As Integer

        Public Declare Function sendto _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer, _
                                         ByVal flags As Integer, _
                                         to_addr As sockaddr, _
                                         ByVal tolen As Integer) As Integer

        Public Declare Function setsockopt _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByVal level As Integer, _
                                         ByVal optname As Integer, _
                                         optval As Any, _
                                         ByVal optlen As Integer) As Integer

        Public Declare Function ShutDown _
                       Lib "ws2_32.DLL" _
                       Alias "shutdown" (ByVal S As Integer, _
                                         ByVal how As Integer) As Integer

        Public Declare Function Socket _
                       Lib "ws2_32.DLL" _
                       Alias "socket" (ByVal af As Integer, _
                                       ByVal s_type As Integer, _
                                       ByVal Protocol As Integer) As Integer

        '---DATABASE FUNCTIONS
        Public Declare Function gethostbyaddr _
                       Lib "ws2_32.DLL" (addr As Long, _
                                         ByVal addr_len As Integer, _
                                         ByVal addr_type As Integer) As Long

        Public Declare Function gethostbyname _
                       Lib "ws2_32.DLL" (ByVal host_name As String) As Long

        Public Declare Function gethostname _
                       Lib "ws2_32.DLL" (ByVal host_name As String, _
                                         ByVal namelen As Integer) As Integer

        Public Declare Function getservbyport _
                       Lib "ws2_32.DLL" (ByVal Port As Integer, _
                                         ByVal proto As String) As Long

        Public Declare Function getservbyname _
                       Lib "ws2_32.DLL" (ByVal serv_name As String, _
                                         ByVal proto As String) As Long

        Public Declare Function getprotobynumber _
                       Lib "ws2_32.DLL" (ByVal proto As Integer) As Long

        Public Declare Function getprotobyname _
                       Lib "ws2_32.DLL" (ByVal proto_name As String) As Long

        '---WINDOWS EXTENSIONS
        Public Declare Function WSAStartup _
                       Lib "ws2_32.DLL" (ByVal wVR As Integer, _
                                         lpWSAD As WSADataType) As Integer

        Public Declare Function WSACleanup Lib "ws2_32.DLL" () As Integer

        Public Declare Sub WSASetLastError Lib "ws2_32.DLL" (ByVal iError As Integer)

        Public Declare Function WSAGetLastError Lib "ws2_32.DLL" () As Integer

        Public Declare Function WSAIsBlocking Lib "ws2_32.DLL" () As Integer

        Public Declare Function WSAUnhookBlockingHook Lib "ws2_32.DLL" () As Integer

        Public Declare Function WSASetBlockingHook _
                       Lib "ws2_32.DLL" (ByVal lpBlockFunc As Long) As Long

        Public Declare Function WSACancelBlockingCall Lib "ws2_32.DLL" () As Integer

        Public Declare Function WSAAsyncGetServByName _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal serv_name As String, _
                                         ByVal proto As String, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSAAsyncGetServByPort _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal Port As Integer, _
                                         ByVal proto As String, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSAAsyncGetProtoByName _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal proto_name As String, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSAAsyncGetProtoByNumber _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal Number As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSAAsyncGetHostByName _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal host_name As String, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSAAsyncGetHostByAddr _
                       Lib "ws2_32.DLL" (ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         addr As Long, _
                                         ByVal addr_len As Integer, _
                                         ByVal addr_type As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer) As Integer

        Public Declare Function WSACancelAsyncRequest _
                       Lib "ws2_32.DLL" (ByVal hAsyncTaskHandle As Integer) As Integer

        Public Declare Function WSAAsyncSelect _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         ByVal hWnd As Integer, _
                                         ByVal wMsg As Integer, _
                                         ByVal lEvent As Long) As Integer

        Public Declare Function WSARecvEx _
                       Lib "ws2_32.DLL" (ByVal S As Integer, _
                                         buf As Any, _
                                         ByVal buflen As Integer, _
                                         ByVal flags As Integer) As Integer
        'Agregado por Maraxus
        Declare Function WSAAccept _
                Lib "ws2_32.DLL" (ByVal S As Integer, _
                                  pSockAddr As sockaddr, _
                                  AddrLen As Integer, _
                                  ByVal lpfnCondition As Long, _
                                  ByVal dwCallbackData As Long) As Integer
    
        Public Const SOMAXCONN As Integer = &H7FFF            ' Agregado por Maraxus

    #ElseIf Win32 Then

        '---Windows System Functions
        Public Declare Function PostMessage _
                       Lib "user32" _
                       Alias "PostMessageA" (ByVal hWnd As Long, _
                                             ByVal wMsg As Long, _
                                             ByVal wParam As Long, _
                                             ByVal lParam As Long) As Long

        Public Declare Sub MemCopy _
                       Lib "kernel32" _
                       Alias "RtlMoveMemory" (Dest As Any, _
                                              Src As Any, _
                                              ByVal cb&)

        Public Declare Function lstrlen _
                       Lib "kernel32" _
                       Alias "lstrlenA" (ByVal lpString As Any) As Long

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
        Public Declare Function accept _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          addr As sockaddr, _
                                          AddrLen As Long) As Long

        Public Declare Function bind _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          addr As sockaddr, _
                                          ByVal namelen As Long) As Long

        Public Declare Function apiclosesocket _
                       Lib "wsock32.dll" _
                       Alias "closesocket" (ByVal S As Long) As Long

        Public Declare Function connect _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          addr As sockaddr, _
                                          ByVal namelen As Long) As Long

        Public Declare Function ioctlsocket _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByVal Cmd As Long, _
                                          argp As Long) As Long

        Public Declare Function getpeername _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          sName As sockaddr, _
                                          namelen As Long) As Long

        Public Declare Function getsockname _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          sName As sockaddr, _
                                          namelen As Long) As Long

        Public Declare Function getsockopt _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByVal level As Long, _
                                          ByVal optname As Long, _
                                          optval As Any, _
                                          optlen As Long) As Long

        Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

        Public Declare Function htons _
                       Lib "wsock32.dll" (ByVal hostshort As Long) As Integer

        Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

        Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long

        Public Declare Function listen _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByVal backlog As Long) As Long

        Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long

        Public Declare Function ntohs _
                       Lib "wsock32.dll" (ByVal netshort As Long) As Integer

        Public Declare Function recv _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByRef buf As Any, _
                                          ByVal buflen As Long, _
                                          ByVal flags As Long) As Long

        Public Declare Function recvfrom _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long, _
                                          ByVal flags As Long, _
                                          from As sockaddr, _
                                          fromlen As Long) As Long

        Public Declare Function ws_select _
                       Lib "wsock32.dll" _
                       Alias "select" (ByVal nfds As Long, _
                                       readfds As fd_set, _
                                       writefds As fd_set, _
                                       exceptfds As fd_set, _
                                       timeout As timeval) As Long

        Public Declare Function send _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long, _
                                          ByVal flags As Long) As Long

        Public Declare Function sendto _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long, _
                                          ByVal flags As Long, _
                                          to_addr As sockaddr, _
                                          ByVal tolen As Long) As Long

        Public Declare Function setsockopt _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByVal level As Long, _
                                          ByVal optname As Long, _
                                          optval As Any, _
                                          ByVal optlen As Long) As Long

        Public Declare Function ShutDown _
                       Lib "wsock32.dll" _
                       Alias "shutdown" (ByVal S As Long, _
                                         ByVal how As Long) As Long

        Public Declare Function Socket _
                       Lib "wsock32.dll" _
                       Alias "socket" (ByVal af As Long, _
                                       ByVal s_type As Long, _
                                       ByVal Protocol As Long) As Long

        '---DATABASE FUNCTIONS
        Public Declare Function gethostbyaddr _
                       Lib "wsock32.dll" (addr As Long, _
                                          ByVal addr_len As Long, _
                                          ByVal addr_type As Long) As Long

        Public Declare Function gethostbyname _
                       Lib "wsock32.dll" (ByVal host_name As String) As Long

        Public Declare Function gethostname _
                       Lib "wsock32.dll" (ByVal host_name As String, _
                                          ByVal namelen As Long) As Long

        Public Declare Function getservbyport _
                       Lib "wsock32.dll" (ByVal Port As Long, _
                                          ByVal proto As String) As Long

        Public Declare Function getservbyname _
                       Lib "wsock32.dll" (ByVal serv_name As String, _
                                          ByVal proto As String) As Long

        Public Declare Function getprotobynumber _
                       Lib "wsock32.dll" (ByVal proto As Long) As Long

        Public Declare Function getprotobyname _
                       Lib "wsock32.dll" (ByVal proto_name As String) As Long

        '---WINDOWS EXTENSIONS
        Public Declare Function WSAStartup _
                       Lib "wsock32.dll" (ByVal wVR As Long, _
                                          lpWSAD As WSADataType) As Long

        Public Declare Function WSACleanup Lib "wsock32.dll" () As Long

        Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)

        Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long

        Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long

        Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long

        Public Declare Function WSASetBlockingHook _
                       Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long

        Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long

        Public Declare Function WSAAsyncGetServByName _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal serv_name As String, _
                                          ByVal proto As String, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSAAsyncGetServByPort _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal Port As Long, _
                                          ByVal proto As String, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSAAsyncGetProtoByName _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal proto_name As String, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSAAsyncGetProtoByNumber _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal Number As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSAAsyncGetHostByName _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal host_name As String, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSAAsyncGetHostByAddr _
                       Lib "wsock32.dll" (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          addr As Long, _
                                          ByVal addr_len As Long, _
                                          ByVal addr_type As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long) As Long

        Public Declare Function WSACancelAsyncRequest _
                       Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long

        Public Declare Function WSAAsyncSelect _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal lEvent As Long) As Long

        Public Declare Function WSARecvEx _
                       Lib "wsock32.dll" (ByVal S As Long, _
                                          buf As Any, _
                                          ByVal buflen As Long, _
                                          ByVal flags As Long) As Long
        'Agregado por Maraxus
        Declare Function WSAAccept _
                Lib "ws2_32.DLL" (ByVal S As Long, _
                                  pSockAddr As sockaddr, _
                                  AddrLen As Long, _
                                  ByVal lpfnCondition As Long, _
                                  ByVal dwCallbackData As Long) As Long

        Public Const SOMAXCONN As Long = &H7FFFFFFF            ' Agregado por Maraxus

    #End If

    'SOME STUFF I ADDED
    Public MySocket%

    Public SockReadBuffer$

    Public Const WSA_NoName = "Unknown"

    Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled

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

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Function ConnectSock(ByVal Host$, _
                         ByVal Port%, _
                         retIpPort$, _
                         ByVal HWndToMsg%, _
                         ByVal Async%) As Integer

        Dim S%, SelectOps%, dummy%

    #ElseIf Win32 Then
        Function ConnectSock(ByVal Host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long

            Dim S&, SelectOps&, dummy&

        #End If

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

#If Win32 Then
    Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    #Else
        Public Function SetSockLinger(ByVal SockNum%, ByVal OnOff%, ByVal LingerTime%) As Integer
        #End If

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
    #If Win32 Then

        Dim nStr&

    #Else

        Dim nStr%

    #End If

    Dim lpStr&

    Dim retString$

    retString = String(32, 0)
    lpStr = inet_ntoa(inn)

    If lpStr Then
        nStr = lstrlen(lpStr)

        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
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
        MemCopy heDestHost, ByVal phe, hostent_size
        HostName = String(256, 0)
        MemCopy ByVal HostName, ByVal heDestHost.h_name, 256
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
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
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

#If Win16 Then
    Public Function GetPeerAddress(ByVal S%) As String

        Dim AddrLen%

    #ElseIf Win32 Then
        Public Function GetPeerAddress(ByVal S&) As String

            Dim AddrLen&

        #End If

        Dim sa As sockaddr

        AddrLen = sockaddr_size

        If getpeername(S, sa, AddrLen) Then
            GetPeerAddress = vbNullString
        Else
            GetPeerAddress = SockAddressToString(sa)

        End If

    End Function

#If Win16 Then
    Public Function GetPortFromString(ByVal PortStr$) As Integer
    #ElseIf Win32 Then
        Public Function GetPortFromString(ByVal PortStr$) As Long
        #End If

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

#If Win16 Then
    Function GetProtocolByName(ByVal Protocol$) As Integer

        Dim tmpShort%

    #ElseIf Win32 Then
        Function GetProtocolByName(ByVal Protocol$) As Long

            Dim tmpShort&

        #End If

        Dim ppe&

        Dim peDestProt As protoent

        ppe = getprotobyname(Protocol)

        If ppe Then
            MemCopy peDestProt, ByVal ppe, protoent_size
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

#If Win16 Then
    Function GetServiceByName(ByVal service$, ByVal Protocol$) As Integer

        Dim Serv%

    #ElseIf Win32 Then
        Function GetServiceByName(ByVal service$, ByVal Protocol$) As Long

            Dim Serv&

        #End If

        Dim pse&

        Dim seDestServ As servent

        pse = getservbyname(service, Protocol)

        If pse Then
            MemCopy seDestServ, ByVal pse, servent_size
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
#If Win16 Then
    Function GetSockAddress(ByVal S%) As String

        Dim AddrLen%

        Dim ret%

    #ElseIf Win32 Then
        Function GetSockAddress(ByVal S&) As String

            Dim AddrLen&

            Dim ret&

        #End If

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

    #If Win16 Then

        Dim nStr%

    #ElseIf Win32 Then

        Dim nStr&

    #End If

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
    MemCopy ByVal retString, ByVal lpStr, nStr
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

'this function DOES work on 16 and 32 bit systems
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
#If Win16 Then
    Public Function ListenForConnect(ByVal Port%, _
                                     ByVal HWndToMsg%, _
                                     ByVal Enlazar As String) As Integer

        Dim S%, dummy%

        Dim SelectOps%

    #ElseIf Win32 Then
        Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, ByVal Enlazar As String) As Long

            Dim S&, dummy&

            Dim SelectOps&

        #End If

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

        '    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
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

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Public Function kSendData(ByVal S%, vMessage As Variant) As Integer
    #ElseIf Win32 Then
        Public Function kSendData(ByVal S&, vMessage As Variant) As Long
        #End If

        Dim TheMsg() As Byte, sTemp$

        TheMsg = vbNullString

        Select Case VarType(vMessage)

            Case 8209   'byte array
                sTemp = vMessage
                TheMsg = sTemp

            Case 8      'string, if we recieve a string, its assumed we are linemode
                #If Win32 Then
                    sTemp = StrConv(vMessage, vbFromUnicode)
                #Else
                    sTemp = vMessage
                #End If

            Case Else
                sTemp = CStr(vMessage)
                #If Win32 Then
                    sTemp = StrConv(vMessage, vbFromUnicode)
                #Else
                    sTemp = vMessage
                #End If

        End Select

        TheMsg = sTemp

        If UBound(TheMsg) > -1 Then
            kSendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)

        End If

    End Function

Public Function SockAddressToString(sa As sockaddr) As String
    SockAddressToString = GetAscIP(sa.sin_addr) & ":" & ntohs(sa.sin_port)

End Function

Public Function StartWinsock(sDescription As String) As Boolean

    Dim StartupData As WSADataType

    If Not WSAStartedUp Then

        'If Not WSAStartup(&H101, StartupData) Then
        If Not WSAStartup(&H202, StartupData) Then  'Use sockets v2.2 instead of 1.1 (Maraxus)
            WSAStartedUp = True
            '            Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
            '            Debug.Print "If wVersion == 257 then everything is kewl"
            '            Debug.Print "szDescription="; StartupData.szDescription
            '            Debug.Print "szSystemStatus="; StartupData.szSystemStatus
            '            Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False

        End If

    End If

    StartWinsock = WSAStartedUp

End Function

Public Function WSAMakeSelectReply(TheEvent%, TheError%) As Long
    WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)

End Function

#End If
