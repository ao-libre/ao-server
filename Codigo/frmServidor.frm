VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command26 
      Caption         =   "Reset Listen"
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
      Left            =   1920
      TabIndex        =   26
      Top             =   6180
      Width           =   1455
   End
   Begin VB.PictureBox picFuera 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4350
      Left            =   120
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   6
      Top             =   120
      Width           =   4590
      Begin VB.VScrollBar VS1 
         Height          =   4335
         LargeChange     =   50
         Left            =   4320
         SmallChange     =   17
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picCont 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   0
         ScaleHeight     =   321
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   7
         Top             =   0
         Width           =   4334
         Begin VB.CommandButton Command27 
            Caption         =   "Debug UserList"
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
            TabIndex        =   27
            Top             =   4440
            Width           =   4095
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Administración"
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
            TabIndex        =   8
            Top             =   4200
            Width           =   4095
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Pausar el servidor"
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
            TabIndex        =   9
            Top             =   3960
            Width           =   4095
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Actualizar npcs.dat"
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
            TabIndex        =   10
            Top             =   3720
            Width           =   4095
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Reload MD5s"
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
            TabIndex        =   25
            Top             =   3480
            Width           =   4095
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Reload Server.ini"
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
            TabIndex        =   11
            Top             =   3240
            Width           =   4095
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Reload Balance.dat"
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
            TabIndex        =   28
            Top             =   3000
            Width           =   4095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Update MOTD"
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
            TabIndex        =   12
            Top             =   2760
            Width           =   4095
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Unban All IPs (PELIGRO!)"
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
            TabIndex        =   13
            Top             =   2520
            Width           =   4095
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Unban All (PELIGRO!)"
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
            TabIndex        =   14
            Top             =   2280
            Width           =   4095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Debug Npcs"
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
            TabIndex        =   15
            Top             =   2040
            Width           =   4095
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Stats de los slots"
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
            TabIndex        =   16
            Top             =   1800
            Width           =   4095
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Trafico"
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
            TabIndex        =   17
            Top             =   1560
            Width           =   4095
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Reload Lista Nombres Prohibidos"
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
            TabIndex        =   18
            Top             =   1320
            Width           =   4095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Actualizar hechizos"
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
            TabIndex        =   19
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Configurar intervalos"
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
            TabIndex        =   20
            Top             =   840
            Width           =   4095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Reiniciar"
            Enabled         =   0   'False
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
            TabIndex        =   21
            Top             =   600
            Width           =   4095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "ReSpawn Guardias en posiciones originales"
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
            TabIndex        =   22
            Top             =   360
            Width           =   4095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Actualizar objetos.dat"
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
            TabIndex        =   23
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Boton Magico para apagar server"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
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
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   4095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guardar todos los personajes"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   6180
      Width           =   945
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset sockets"
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
      Left            =   240
      TabIndex        =   4
      Top             =   6180
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   4560
      Width           =   4335
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Sub Command1_Click()
Call ResetForums
Call LoadOBJData

End Sub

Private Sub Command10_Click()
frmTrafic.Show
End Sub

Private Sub Command11_Click()
frmConID.Show
End Sub

Private Sub Command12_Click()
frmDebugNpc.Show
End Sub

Private Sub Command14_Click()
Call LoadMotd
End Sub

Private Sub Command15_Click()
On Error Resume Next

Dim Fn As String
Dim cad$
Dim N As Integer, k As Integer

Dim sENtrada As String

sENtrada = InputBox("Escribe ""estoy DE acuerdo"" entre comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes.", "UnBan", "hola")
If sENtrada = "estoy DE acuerdo" Then

    Fn = App.Path & "\logs\GenteBanned.log"
    
    If FileExist(Fn, vbNormal) Then
        N = FreeFile
        Open Fn For Input Shared As #N
        Do While Not EOF(N)
            k = k + 1
            Input #N, cad$
            Call UnBan(cad$)
            
        Loop
        Close #N
        MsgBox "Se han habilitado " & k & " personajes."
        Kill Fn
    End If
End If

End Sub

Private Sub Command16_Click()
Call LoadSini
End Sub

Private Sub Command17_Click()
    Call CargaNpcsDat
End Sub

Private Sub Command18_Click()
Me.MousePointer = 11
Call mdParty.ActualizaExperiencias
Call GuardarUsuarios
Me.MousePointer = 0
MsgBox "Grabado de personajes OK!"
End Sub

Private Sub Command19_Click()
Dim i As Long, N As Long

Dim sENtrada As String

sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes", "UnBan", "hola")
If sENtrada = "estoy DE acuerdo" Then
    
    N = BanIps.Count
    For i = 1 To BanIps.Count
        BanIps.Remove 1
    Next i
    
    MsgBox "Se han habilitado " & N & " ipes"
End If

End Sub

Private Sub Command2_Click()
frmServidor.Visible = False
End Sub

Private Sub Command20_Click()
#If UsarQueSocket = 1 Then

If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    Call WSApiReiniciarSockets
End If

#ElseIf UsarQueSocket = 2 Then

Dim LoopC As Integer

If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnID <> -1 And UserList(LoopC).ConnIDValida Then
            Call CloseSocket(LoopC)
        End If
    Next LoopC
    
    Call frmMain.Serv.Detener
    Call frmMain.Serv.Iniciar(Puerto)
End If

#End If
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()

If EnPausa = False Then
    EnPausa = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Command21.Caption = "Reanudar el servidor"
Else
    EnPausa = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Command21.Caption = "Pausar el servidor"
End If

End Sub

Private Sub Command22_Click()
    Me.Visible = False
    frmAdmin.Show
End Sub

Private Sub Command23_Click()
If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, "Apagar Magicamente") = vbYes Then
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
End If
End Sub

Private Sub Command25_Click()
Call MD5sCarga

End Sub

Private Sub Command26_Click()
#If UsarQueSocket = 1 Then
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
#End If
End Sub

Private Sub Command27_Click()
frmUserList.Show

End Sub

Private Sub Command28_Click()
    Call LoadBalance
End Sub

Private Sub Command3_Click()
If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la pérdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
    Me.Visible = False
    Call General.Restart
End If
End Sub

Private Sub Command4_Click()
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call ES.DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"


#If UsarQueSocket = 1 Then
Call apiclosesocket(SockListen)
#ElseIf UsarQueSocket = 0 Then
frmMain.Socket1.Cleanup
frmMain.Socket2(0).Cleanup
#ElseIf UsarQueSocket = 2 Then
frmMain.Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

Call FreeNPCs
Call FreeCharIndexes

Call LoadSini
Call CargarBackUp
Call LoadOBJData

#If UsarQueSocket = 1 Then
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = Puerto
frmMain.Socket1.listen
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
FrmInterv.Show
End Sub

Private Sub Command8_Click()
Call CargarHechizos
End Sub

Private Sub Command9_Click()
Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
frmServidor.Visible = False
End Sub

Private Sub Form_Load()
#If UsarQueSocket = 1 Then
Command20.Visible = True
Command26.Visible = True
#ElseIf UsarQueSocket = 0 Then
Command20.Visible = False
Command26.Visible = False
#ElseIf UsarQueSocket = 2 Then
Command20.Visible = True
Command26.Visible = False
#End If

VS1.min = 0
If picCont.Height > picFuera.ScaleHeight Then
    VS1.max = picCont.Height - picFuera.ScaleHeight
Else
    VS1.max = 0
End If
picCont.Top = -VS1.value

End Sub

Private Sub VS1_Change()
picCont.Top = -VS1.value
End Sub

Private Sub VS1_Scroll()
picCont.Top = -VS1.value
End Sub
