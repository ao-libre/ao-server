VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.1#0"; "MSWINSCK.OCX"
Begin VB.Form frmBotAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BotAPI Server"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   0
      RemoteHost      =   "localhost"
      RemotePort      =   0
   End
End
Attribute VB_Name = "frmBotAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' frmBotAPI.frm â€” BotAPI TCP listener form.
'
' Socket(0) is the listen socket on port BOTAPI_PORT.
' Socket(1..MAX_BOT_CONNECTIONS) are accepted client sockets.
' Each client sends newline-delimited JSON; responses are sent back with \n.
''

Option Explicit

Private Const MAX_BOT_CONNECTIONS As Integer = 20

' Per-connection receive buffers (one per socket index)
Private m_buffer(0 To MAX_BOT_CONNECTIONS) As String

' ---- Start listening -------------------------------------------------------
Public Sub StartListening()

    On Error GoTo ErrHandler

    Socket(0).LocalPort = BOTAPI_PORT
    Socket(0).Listen

    Call LogStatic("BotAPI: Listening on port " & BOTAPI_PORT)

    Exit Sub

ErrHandler:
    Call LogStatic("BotAPI: Failed to start listener on port " & BOTAPI_PORT & " - " & Err.Description)

End Sub

' ---- Socket(0): new connection request ------------------------------------
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Index <> 0 Then Exit Sub

    ' Find a free slot
    Dim slot As Integer
    For slot = 1 To MAX_BOT_CONNECTIONS
        If Socket(slot).State = sckClosed Then
            Load Socket(slot)
            m_buffer(slot) = ""
            Socket(slot).Accept requestID
            Call LogStatic("BotAPI: Connection accepted on slot " & slot)
            Exit Sub
        End If
    Next slot

    ' No free slot â€” reject by not accepting (connection will time out)
    Call LogStatic("BotAPI: All slots full, connection rejected")

End Sub

' ---- Socket(1..N): data received ------------------------------------------
Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If Index = 0 Then Exit Sub

    Dim chunk As String
    Socket(Index).GetData chunk, vbString

    m_buffer(Index) = m_buffer(Index) & chunk

    ' Process all complete newline-delimited messages
    Dim pos As Integer
    pos = InStr(m_buffer(Index), Chr(10))
    Do While pos > 0

        Dim line As String
        line = Left$(m_buffer(Index), pos - 1)
        ' Strip CR if present
        If Len(line) > 0 And Right$(line, 1) = Chr(13) Then
            line = Left$(line, Len(line) - 1)
        End If
        m_buffer(Index) = Mid$(m_buffer(Index), pos + 1)

        line = Trim$(line)
        If Len(line) > 0 Then
            Dim response As String
            response = BotDispatch(line)
            On Error Resume Next
            Socket(Index).SendData response & Chr(10)
            On Error GoTo 0
        End If

        pos = InStr(m_buffer(Index), Chr(10))

    Loop

End Sub

' ---- Socket(1..N): closed --------------------------------------------------
Private Sub Socket_Close(Index As Integer)

    If Index = 0 Then Exit Sub

    m_buffer(Index) = ""
    On Error Resume Next
    Socket(Index).Close
    Unload Socket(Index)
    On Error GoTo 0

    Call LogStatic("BotAPI: Slot " & Index & " disconnected")

End Sub

' ---- Socket(1..N): error ---------------------------------------------------
Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, _
                         Description As String, ByVal Scode As Long, _
                         ByVal Source As String, ByVal HelpFile As String, _
                         ByVal HelpContext As Long, CancelDisplay As Boolean)

    CancelDisplay = True
    m_buffer(Index) = ""

    On Error Resume Next
    If Index = 0 Then
        Socket(0).Close
        Socket(0).Listen
    Else
        Socket(Index).Close
        Unload Socket(Index)
    End If
    On Error GoTo 0

    Call LogStatic("BotAPI: Socket error on slot " & Index & ": " & Description)

End Sub
