VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   6390
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox chkServerHabilitado 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Server Habilitado Solo Gms"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtNumUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSystray 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Systray"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrarServer 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cerrar Servidor"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfiguracion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configuraci�n General"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   4935
   End
   Begin VB.CommandButton cmdDump 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Log Cr�tico de Usuarios"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   4935
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   360
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mensajea todos los clientes (Solo testeo)"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.TextBox txtChat 
         BackColor       =   &H00C0FFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label Escuch 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N�mero de usuarios jugando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
        With UserList(iUserIndex)
            'Conexion activa? y es un usuario loggeado?
            If .ConnID <> -1 And .flags.UserLogged Then
                'Actualiza el contador de inactividad
                If .flags.Traveling = 0 Then
                    .Counters.IdleCount = .Counters.IdleCount + 1
                End If
                
                If .Counters.IdleCount >= IdleLimit Then
                    Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")
                    'mato los comercios seguros
                    If .ComUsu.DestUsu > 0 Then
                        If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                                Call FinComerciarUsu(.ComUsu.DestUsu)
                                Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away
                            End If
                        End If
                        Call FinComerciarUsu(iUserIndex)
                    End If
                    Call Cerrar_Usuario(iUserIndex)
                End If
            End If
        End With
    Next iUserIndex
End Sub


Private Sub AutoSave_Timer()

On Error GoTo ErrHandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long


Minutos = Minutos + 1
MinsPjesSave = MinsPjesSave + 1

'�?�?�?�?�?�?�?�?�?�?�
Call ModAreas.AreasOptimizacion
'�?�?�?�?�?�?�?�?�?�?�

'Actualizamos la lluvia
Call tLluviaEvent

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
    KillLog
End If

If Minutos >= MinutosWs Then
    Call ES.DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinsPjesSave = MinutosGuardarUsuarios - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CharSave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
ElseIf MinsPjesSave >= MinutosGuardarUsuarios Then
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    MinsPjesSave = 0
End If

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
ErrHandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)
    Resume Next
End Sub


Private Sub chkServerHabilitado_Click()
    ServerSoloGMs = chkServerHabilitado.value
End Sub

Private Sub cmdCerrarServer_Click()
    If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. " & _
        "�Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        
        prgRun = False
        Dim f
        For Each f In Forms
            Unload f
        Next
    End If
End Sub

Private Sub cmdConfiguracion_Click()
    frmServidor.Visible = True
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

    Dim i As Integer
    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & _
            ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & _
            " UserLogged: " & UserList(i).flags.UserLogged)
    Next i
    
    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub cmdSystray_Click()
    SetSystray
End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call SetSystray
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call Statistics.DumpStatistics

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server cerrado."
Close #N

End

End Sub


Private Sub mnusalir_Click()
    Call cmdCerrarServer_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If

End Sub

Private Sub SetSystray()

    Dim i As Integer
    Dim S As String
    Dim nid As NOTIFYICONDATA
    
    S = "ARGENTUM-ONLINE"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
        
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

End Sub

Private Sub tLluviaEvent()
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = True
            MinutosSinLluvia = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    ElseIf MinutosSinLluvia >= 1440 Then
        Lloviendo = True
        MinutosSinLluvia = 0
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
        Lloviendo = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        MinutosLloviendo = 0
    Else
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = False
            MinutosLloviendo = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    End If
End If
End Sub







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.description)

End Sub





#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txtChat_Change()

End Sub
