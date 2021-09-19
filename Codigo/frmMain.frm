VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   6975
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   10425
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
   ScaleHeight     =   6975
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtRecordOnline 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1575
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frmMain.frx":1042
      Top             =   4800
      Width           =   4935
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9600
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdForzarCierre 
      BackColor       =   &H008080FF&
      Caption         =   "Forzar Cierre del Servidor Sin Backup"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   4935
   End
   Begin VB.CheckBox chkServerHabilitado 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Server Habilitado Solo Gms"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtNumUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
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
      TabIndex        =   9
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdApagarServidor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apagar Servidor Con Backup"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfiguracion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configuracion General"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   4935
   End
   Begin VB.CommandButton cmdDump 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Log Critico de Usuarios"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
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
      TabIndex        =   1
      Top             =   720
      Width           =   4935
      Begin VB.Timer TimerEnviarDatosServer 
         Interval        =   65535
         Left            =   2760
         Top             =   1440
      End
      Begin VB.Timer GameTimer 
         Interval        =   40
         Left            =   2160
         Top             =   1440
      End
      Begin VB.Timer TIMER_AI 
         Interval        =   380
         Left            =   1680
         Top             =   1440
      End
      Begin VB.Timer PacketResend 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1200
         Top             =   1440
      End
      Begin VB.Timer Auditoria 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   1440
      End
      Begin VB.TextBox txtChat 
         BackColor       =   &H00C0FFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label lblLloviendoInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Esta lloviendo? Cargando..."
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   4440
      Width           =   4455
   End
   Begin VB.Label lblRespawnNpcs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Tiempo restante para Respawn Npc : Cargando..."
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label lblCharSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Tiempo restante para Char Save : Cargando..."
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label lblWorldSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Tiempo Restante para World Save: Cargando..."
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label lblFooter 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "http://www.ArgentumOnline.org"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Label lblIpHelpText 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hace click sobre tu ip para copiarla"
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
      Height          =   225
      Left            =   5400
      TabIndex        =   19
      Top             =   2880
      Width           =   2970
   End
   Begin VB.Label lblRecordOnline 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Record usuarios online:"
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
      Left            =   7080
      TabIndex        =   18
      Top             =   360
      Width           =   1965
   End
   Begin VB.Label lblIpTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP:PUERTO - Comparti esta informacion a quien quieras que se conecte a tu servidor."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label lblIp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Click para revelar)"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Escuch 
      BackColor       =   &H80000017&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios jugando:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   2460
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
'Copyright (C) 2002 Marquez Pablo Ignacio
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

Public WithEvents WinsockThread As clsSubclass
Attribute WinsockThread.VB_VarHelpID = -1

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

Const CANTIDAD_MINUTOS_NUEVA_LLUVIA_EN_JUEGO = 65

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hWnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function Shell_NotifyIconA _
                Lib "SHELL32" (ByVal dwMessage As Long, _
                               lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, _
                                   ID As Long, _
                                   flags As Long, _
                                   CallbackMessage As Long, _
                                   Icon As Long, _
                                   Tip As String) As NOTIFYICONDATA

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
                
                If Not EsGm(iUserIndex) Then
                    If .Counters.IdleCount >= IdleLimit Then
                        Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")

                        'mato los comercios seguros
                        If .ComUsu.DestUsu > 0 Then
                            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                                    Call FinComerciarUsu(.ComUsu.DestUsu)

                                End If

                            End If

                            Call FinComerciarUsu(iUserIndex)

                        End If

                        Call Cerrar_Usuario(iUserIndex)

                    End If

                End If

            End If

        End With

    Next iUserIndex

End Sub

Public Sub UpdateNpcsExp(ByVal Multiplicador As Single) ' 0.13.5
    Dim NpcIndex As Long
    For NpcIndex = 1 To LastNPC
        With Npclist(NpcIndex)
            .GiveEXP = .GiveEXP * Multiplicador
            .flags.ExpCount = .flags.ExpCount * Multiplicador
        End With
    Next NpcIndex
End Sub

Private Sub HappyHourManager()
    If iniHappyHourActivado = True Then
        Dim tmpHappyHour As Double
    
        ' HappyHour
        Dim iDay As Integer ' 0.13.5
        Dim Message As String

        iDay = Weekday(Date)
        tmpHappyHour = HappyHourDays(iDay).Multi
         
        If tmpHappyHour <> HappyHour Then ' 0.13.5
            If HappyHourActivated Then
                ' Reestablece la exp de los npcs
                If HappyHour <> 0 Then Call UpdateNpcsExp(1 / HappyHour)
            End If
           
            If tmpHappyHour = 1 Then ' Desactiva
                Message = "Ha concluido la Happy Hour!"
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_SERVER))
                HappyHourActivated = False

                If ConexionAPI Then
                    Call ApiEndpointSendHappyHourEndedMessageDiscord(Message)
                End If
           
            Else ' Activa?
                If HappyHourDays(iDay).Hour = Hour(Now) And tmpHappyHour > 0 Then ' GSZAO - Es la hora pautada?
                    UpdateNpcsExp tmpHappyHour
                    
                    If HappyHour <> 1 Then
                        Message = "Se ha modificado la Happy Hour, a partir de ahora las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%"
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_SERVER))

                        If ConexionAPI Then
                            Call ApiEndpointSendHappyHourModifiedMessageDiscord(Message)
                        End If
                    Else
                        Message = "Ha comenzado la Happy Hour! Las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%!"

                       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_SERVER))
                    
                        'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
                        'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
                        'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
                        If ConexionAPI Then
                            Call ApiEndpointSendHappyHourStartedMessageDiscord(Message)
                        End If

                    End If
                    
                    HappyHourActivated = True
                Else
                    HappyHourActivated = False ' GSZAO
                End If
            End If
         
            HappyHour = tmpHappyHour
        End If
    Else
        ' Si estaba activado, lo deshabilitamos
        If HappyHour <> 0 Then
            Call UpdateNpcsExp(1 / HappyHour)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_SERVER))
            HappyHourActivated = False
            HappyHour = 0
        End If
    End If
End Sub

Private Sub Auditoria_Timer()
    Call mMainLoop.Auditoria
End Sub

Private Sub AutoSave_Timer()

    On Error GoTo errHandler

    'fired every minute
    Static Minutos          As Long

    Static MinutosLatsClean As Long

    Static MinsPjesSave     As Long

    Static MinsEventoPesca  As Long

    MinsEventoPesca = MinsEventoPesca + 1
    Minutos = Minutos + 1
    MinsPjesSave = MinsPjesSave + 1

    Call HappyHourManager
    
    'Actualizamos el Centinela en caso de que este activo en el server.ini
    If isCentinelaActivated Then
        Call modCentinela.ChekearUsuarios
    End If

    'Actualizamos la lluvia
    Call tLluviaEvent

    If Minutos = MinutosWs - 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_SERVER))
        KillLog

    ElseIf Minutos >= MinutosWs Then
        Call ES.DoBackUp
        Call aClon.VaciarColeccion
        Minutos = 0

    End If

    If MinsPjesSave = MinutosGuardarUsuarios - 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CharSave en 1 minuto ...", FontTypeNames.FONTTYPE_SERVER))
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

    Call CheckEstadoDelMar(MinsEventoPesca)

    Call CheckIdleUser

    frmMain.lblWorldSave.Caption = "Proximo WorldSave: " & MinutosWs - Minutos & " Minutos"
    frmMain.lblCharSave.Caption = "Proximo CharSave: " & MinutosGuardarUsuarios - MinsPjesSave & " Minutos"
    frmMain.lblRespawnNpcs.Caption = "Respawn Npcs a POS originales: " & 15 - MinutosLatsClean & " Minutos"

    '<<<<<-------- Log the number of users online ------>>>
    Dim n As Integer

    n = FreeFile()
    Open App.Path & "\logs\numusers.log" For Output Shared As n
    Print #n, NumUsers
    Close #n
    '<<<<<-------- Log the number of users online ------>>>

    Exit Sub

errHandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)

    Resume Next

End Sub

Private Sub chkServerHabilitado_Click()
    ServerSoloGMs = chkServerHabilitado.Value

End Sub

Private Sub cmdApagarServidor_Click()

    If MsgBox("Realmente desea cerrar el servidor?", vbYesNo, "CIERRE DEL SERVIDOR!!!") = vbNo Then Exit Sub
    
    Me.MousePointer = 11
    
    FrmStat.Show
    
    'WorldSave
    Call ES.DoBackUp

    'commit experiencia
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios

    'Chauuu
    Unload frmMain

    Call CloseServer
    
End Sub

Private Sub cmdConfiguracion_Click()
    frmServidor.Visible = True

End Sub

Private Sub CMDDUMP_Click()

    On Error Resume Next

    Dim i As Integer

    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
    Next i
    
    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub cmdForzarCierre_Click()
        
    If MsgBox("Desea FORZAR el CIERRE del SERVIDOR?", vbYesNo, "CIERRE DEL SERVIDOR!!!") = vbNo Then Exit Sub
        
    Call CloseServer

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

                If hHook Then
                    UnhookWindowsHookEx hHook
                    hHook = 0
                End If


        End Select

    End If
   
End Sub

Private Sub QuitarIconoSystray()

    On Error Resume Next

    'Borramos el icono del systray
    Dim i   As Integer

    Dim nid As NOTIFYICONDATA

    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

    i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    'Save stats!!!
    Call Statistics.DumpStatistics

    Call QuitarIconoSystray

    Call LimpiaWsApi

    Dim LoopC As Integer
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
    Next

    'Log
    Dim n As Integer: n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
        Print #n, Date & " " & time & " server cerrado."
    Close #n

    End

End Sub

Private Sub GameTimer_Timer()
    Call mMainLoop.GameTimer
End Sub

Private Sub lblIp_Click()
    Clipboard.Clear
    Clipboard.SetText (lblIp.Caption)
    
    If lblIp.Caption = lblIp.Tag Then
        lblIp.Caption = lblIp.Tag
    Else
        lblIp.Caption = "(Click para revelar)"
    End If
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " | La ip y puerto fueron copiadas correctamente, pegalas donde quieras."
End Sub

Private Sub mnusalir_Click()
    Call cmdApagarServidor_Click

End Sub

Public Sub mnuMostrar_Click()

    On Error Resume Next

    WindowState = vbNormal
    Call Form_MouseMove(0, 0, 7725, 0)

End Sub

Private Sub KillLog()

    On Error Resume Next

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"

    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then
            Kill App.Path & "\logs\wsapi.log"
        End If
    End If

End Sub

Private Sub SetSystray()

    Dim i   As Integer

    Dim S   As String

    Dim nid As NOTIFYICONDATA
    
    S = "ARGENTUM ONLINE LIBRE - http://www.ArgentumOnline.org"
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
        
        If MinutosSinLluvia >= CANTIDAD_MINUTOS_NUEVA_LLUVIA_EN_JUEGO Then
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

Private Sub PacketResend_Timer()
    Call mMainLoop.PacketResend
End Sub

Private Sub TIMER_AI_Timer()
    Call mMainLoop.TIMER_AI
End Sub

Private Sub TimerEnviarDatosServer_Timer()
    Call mMainLoop.TimerEnviarDatosServer
End Sub

Public Sub WinsockThread_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long, DefCall As Boolean)

    On Error Resume Next

    Dim Ret      As Long
    Dim Tmp()    As Byte
    Dim S        As Long
    Dim e        As Long
    Dim n        As Integer
    Dim UltError As Long
    
    Select Case Msg

        Case 1025
            S = wParam
            e = WSAGetSelectEvent(lParam)
            
            Select Case e

                Case FD_ACCEPT
                    If S = SockListen Then
                        Call EventoSockAccept(S)
                    End If

                Case FD_READ
                    n = BuscaSlotSock(S)

                    If n < 0 And S <> SockListen Then
                        Call WSApiCloseSocket(S)
                        Exit Sub
                    End If
                    
                    'create appropiate sized buffer
                    ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
                    
                    Ret = recv(S, Tmp(0), SIZE_RCVBUF, 0)

                    ' Comparo por = 0 ya que esto es cuando se cierra
                    ' "gracefully". (mas abajo)
                    If Ret < 0 Then
                        UltError = Err.LastDllError

                        If UltError = WSAEMSGSIZE Then
                            Debug.Print "WSAEMSGSIZE"
                            Ret = SIZE_RCVBUF
                        
                        Else
                            Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                            Call LogApiSock("Error en Recv: N=" & n & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                            
                            'no hay q llamar a CloseSocket() directamente,
                            'ya q pueden abusar de algun error para
                            'desconectarse sin los 10segs. CREEME.
                            Call CloseSocketSL(n)
                            Call Cerrar_Usuario(n)
                            Exit Sub

                        End If

                    ElseIf Ret = 0 Then
                        Call CloseSocketSL(n)
                        Call Cerrar_Usuario(n)

                    End If
                    
                    ReDim Preserve Tmp(Ret - 1) As Byte
                    
                    Call EventoSockRead(n, Tmp)
                
                Case FD_CLOSE
                    n = BuscaSlotSock(S)

                    If S <> SockListen Then Call apiclosesocket(S)
                    
                    If n > 0 Then
                        Call BorraSlotSock(S)
                        UserList(n).ConnID = -1
                        UserList(n).ConnIDValida = False
                        Call EventoSockClose(n)
                    End If

            End Select
        
        Case Else
            DefCall = True

    End Select

End Sub
