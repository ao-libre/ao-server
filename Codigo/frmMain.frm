VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   1785
   ClientLeft      =   1950
   ClientTop       =   1815
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
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   945
      Top             =   540
   End
   Begin VB.Timer CmdExec 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   60
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   1035
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1020
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1920
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
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
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
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
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

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
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
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
Dim iUserIndex As Integer

For iUserIndex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
   If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            'mato los comercios seguros
            If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(iUserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call FinComerciarUsu(iUserIndex)
            End If
            Call Cerrar_Usuario(iUserIndex)
        End If
  End If
  
Next iUserIndex

End Sub



Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb
Call ActualizaStatsES



Exit Sub

errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsSocketReset As Long
Static MinsPjesSave As Long
Static MinutosNumUsersCheck As Long

Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        
        Horas = 0
        
    End If
    MinsRunning = 0
End If

    
Minutos = Minutos + 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual está.
' alejo.
MinsSocketReset = MinsSocketReset + 1
' cada 1 minutos hacer el checkeo
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
            If UserList(i).Counters.IdleCount > ((IntervaloCerrarConexion * 2) / 3) Then
                Call CloseSocket(i)
            End If
        End If
    Next i
    'Call ReloadSokcet
    
    Call LogCriticEvent("NumUsers: " & NumUsers & " WSAPISock2Usr: " & WSAPISock2Usr.Count)
End If
#End If

MinutosNumUsersCheck = MinutosNumUsersCheck + 1

If MinutosNumUsersCheck >= 2 Then
    MinutosNumUsersCheck = 0
    num = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
            num = num + 1
        End If
    Next i
    If num <> NumUsers Then
        NumUsers = num
        'Call SendData(SendTarget.ToAdmins, 0, 0, "Servidor> Error en NumUsers. Contactar a algun Programador." & FONTTYPE_SERVER)
        Call LogCriticEvent("Num <> NumUsers")
    End If
End If

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, 0, "||Worldsave en 1 minuto ..." & FONTTYPE_VENENO)
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)

End Sub






Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub CmdExec_Timer()
Dim i As Integer
Static N As Long

On Error Resume Next ':(((

N = N + 1

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
        If N >= 10 Then
            If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
    #If UsarQueSocket = 1 Then
                Call IntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 0 Then
    '            Call WrchIntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 2 Then
    '            Call ServIntentarEnviarDatosEncolados(i)
    #ElseIf UsarQueSocket = 3 Then
        'NADA, el control deberia ocuparse de esto!!!
        'si la cola se llena, dispara un on close
    #End If
            End If
        End If
    End If
Next i

If N >= 10 Then
    N = 0
End If

Exit Sub
hayerror:

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> " & BroadMsg.Text & FONTTYPE_SERVER)
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

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Call DescargaNpcsDat


Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer

Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long

'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If

If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If

If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]

On Error GoTo hayerror

 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '¿User valido?

      If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
         
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         
         UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0

         
         Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)
         
                
         If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
         If UserList(iUserIndex).flags.Ceguera = 1 Or _
            UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
         
          
         If UserList(iUserIndex).flags.Muerto = 0 Then
               
               '[Consejeros]
               If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoFrio(iUserIndex)
               If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
               If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoVeneno(iUserIndex, bEnviarStats)
               If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
               If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                
               Call DuracionPociones(iUserIndex)
                
               Call HambreYSed(iUserIndex, bEnviarAyS)
                
               If Lloviendo Then
                    If Not Intemperie(iUserIndex) Then
                        If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                        'No esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                            
                        ElseIf UserList(iUserIndex).flags.Descansar Then
                        'esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                                 'termina de descansar automaticamente
                            If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                    Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                    Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                    UserList(iUserIndex).flags.Descansar = False
                            End If
                            
                        End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    'No esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                        
                    ElseIf UserList(iUserIndex).flags.Descansar Then
                    'esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                             'termina de descansar automaticamente
                        If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                            UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                UserList(iUserIndex).flags.Descansar = False
                        End If
                        
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
               
               If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)

               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
     Else 'no esta logeado?
     'UserList(iUserIndex).Counters.IdleCount = 0
     '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
     'de dados :)
        
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call CloseSocket(iUserIndex)
        End If
        
     End If 'UserLogged

   End If

   Next iUserIndex

'[Alejo]
If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If

If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If

If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If

Exit Sub
hayerror:
LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
'[/Alejo]
  'DoEvents
End Sub

Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
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

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            e_p = esPretoriano(NpcIndex)
            If e_p > 0 Then
                If Npclist(NpcIndex).flags.Paralizado = 1 Then Call EfectoParalisisNpc(NpcIndex)
                Select Case e_p
                    Case 1  ''clerigo
                        Call PRCLER_AI(NpcIndex)
                    Case 2  ''mago
                        Call PRMAGO_AI(NpcIndex)
                    Case 3  ''cazador
                        Call PRCAZA_AI(NpcIndex)
                    Case 4  ''rey
                        Call PRREY_AI(NpcIndex)
                    Case 5  ''guerre
                        Call PRGUER_AI(NpcIndex)
                End Select
            Else
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
            End If
        End If
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i

End Sub

Private Sub tLluvia_Timer()
On Error GoTo errhandler

Dim iCount As Integer
If Lloviendo Then
   For iCount = 1 To LastUser
    Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
errhandler:
Call LogError("tLluvia " & Err.Number & ": " & Err.Description)
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 10 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, 0, "LLU")
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, 0, "LLU")
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, 0, "LLU")
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 7 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(SendTarget.ToAll, 0, 0, "LLU")
            End If
    End If
End If


Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")



End Sub


Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
Static Segundos As Integer
Dim NuevaA As Boolean
Dim NuevoL As Boolean
Dim GI As Integer

Segundos = Segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
            
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    Call SendData(SendTarget.ToIndex, i, 0, "||Estas obstruyendo la via publica, muévete o seras encarcelado!!!" & FONTTYPE_INFO)
                    If UserList(i).Counters.PiqueteC > 23 Then
                            UserList(i).Counters.PiqueteC = 0
                            Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
            Else
                    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If

            'ustedes se preguntaran que hace esto aca?
            'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
            'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
            'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable

            GI = UserList(i).GuildIndex
            If GI > 0 Then
                NuevaA = False
                NuevoL = False
                If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
                    Call SendData(SendTarget.ToIndex, i, 0, "||Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!" & FONTTYPE_GUILD)
                End If
                If NuevaA Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||¡El clan ha pasado a tener alineación neutral!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan cambio de alineacion!")
                End If
                If NuevoL Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||¡El clan tiene un nuevo líder!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan tiene nuevo lider!")
                End If
            End If

            If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
            
    End If
Next i

If Segundos >= 18 Then Segundos = 0
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
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
Call LogError("TCPServer::NuevaConexion " & Err.Description)
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
Dim T() As String
Dim LoopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If

RD = StrConv(Datos, vbUnicode)

'call logindex(MiDato, "Read. ConnId: " & ID & " Midato: " & MiDato & " Dato: " & RD)

UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

T = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(T) > 0 Then
    UserList(MiDato).RDBuffer = T(UBound(T))
    
    For LoopC = 0 To UBound(T) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If T(LoopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(T(LoopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, T(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
