VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Configuracion del Servidor"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listDats 
      Height          =   1425
      Left            =   2160
      TabIndex        =   23
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdReiniciar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Administracion"
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6375
      Begin VB.CommandButton cmdRecargarGuardiasPosOrig 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardias en pos original"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1860
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetListen 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reset Listen"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetSockets 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reset sockets"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdDebugUserlist 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Debug UserList"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdUnbanAllIps 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unban All IPs (PELIGRO!)"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdUnbanAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unban All (PELIGRO!)"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdDebugNpcs 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Debug Npcs"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton frmAdministracion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Administracion"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPausarServidor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pausar el servidor"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdStatsSlots 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Stats de Slots"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdVerTrafico 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Trafico"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdConfigIntervalos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Config. Intervalos"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir (Esc)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdForzarCierre 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Forzar Cierre del Servidor"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Backup"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   6375
      Begin VB.CommandButton cmdLoadWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cargar Mapas"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCharBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Chars"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Mapas"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recargar"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdRecargarAdministradores 
         BackColor       =   &H0080C0FF&
         Caption         =   "Administradores"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdRecargarClanes 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clanes"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1290
         Width           =   1575
      End
      Begin VB.CommandButton cmdRecargarServerIni 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Server.ini"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmServidor"
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

Private Sub cmdRecargarClanes_Click()
    Call LoadGuildsDB
End Sub

Private Sub Form_Load()

    cmdResetSockets.Visible = True
    cmdResetListen.Visible = True
    
    'Listamos el contenido de la carpeta Dats
    Dim sFilename As String
        sFilename = Dir$(DatPath)
    
    Do While sFilename > vbNullString
    
      Call listDats.AddItem(sFilename)
      sFilename = Dir$()
    
    Loop
  
End Sub

Private Sub cmdForzarCierre_Click()
        
    If MsgBox("Desea FORZAR el CIERRE del SERVIDOR?", vbYesNo, "CIERRE DEL SERVIDOR!!!") = vbNo Then Exit Sub
        
    Call CloseServer

End Sub

Private Sub cmdCerrar_Click()
    frmServidor.Visible = False

End Sub

Private Sub cmdCharBackup_Click()
    Me.MousePointer = 11
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Me.MousePointer = 0
    MsgBox "Grabado de personajes OK!"

End Sub

Private Sub cmdConfigIntervalos_Click()
    FrmInterv.Show

End Sub

Private Sub cmdDebugNpcs_Click()
    frmDebugNpc.Show

End Sub

Private Sub cmdDebugUserlist_Click()
    frmUserList.Show

End Sub

Private Sub cmdLoadWorldBackup_Click()

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txtStatus.Text = "Reiniciando."
    
    FrmStat.Show
    
    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

    Call apiclosesocket(SockListen)

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

    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Reiniciando Terminado. Escuchando conexiones entrantes ..."
End Sub

Private Sub cmdPausarServidor_Click()

    If EnPausa = False Then
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Pausar el servidor"

    End If

End Sub

Private Sub cmdRecargarServerIni_Click()
    Call LoadSini
End Sub

Private Sub cmdReiniciar_Click()

    If MsgBox("Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. " & "Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbNo Then Exit Sub
    
    Me.Visible = False
    Call General.Restart

End Sub

Private Sub cmdResetListen_Click()

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

End Sub

Private Sub cmdResetSockets_Click()

    If MsgBox("Esta seguro que desea reiniciar los sockets? Se cerraran todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
        Call WSApiReiniciarSockets
    End If

End Sub

Private Sub cmdStatsSlots_Click()
    frmConID.Show

End Sub

Private Sub cmdUnbanAll_Click()

    On Error Resume Next

    Dim Fn       As String
    Dim cad$
    Dim n        As Integer, K As Integer
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" entre comillas y con distincion de mayusculas minusculas para desbanear a todos los personajes.", "UnBan", "hola")

    If sENtrada = "estoy DE acuerdo" Then
    
        Fn = App.Path & "\logs\GenteBanned.log"
        
        If FileExist(Fn, vbNormal) Then
            n = FreeFile
            Open Fn For Input Shared As #n

            Do While Not EOF(n)
                K = K + 1
                Input #n, cad$
                Call UnBan(cad$)
                
            Loop
            Close #n
            MsgBox "Se han habilitado " & K & " personajes."
            Kill Fn

        End If

    End If

End Sub

Private Sub cmdUnbanAllIps_Click()

    Dim i        As Long, n As Long
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distincion de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")

    If sENtrada = "estoy DE acuerdo" Then
        
        n = BanIps.Count

        For i = 1 To BanIps.Count
            Call BanIpQuita(BanIps(i))
        Next i
        
        MsgBox "Se han habilitado " & n & " ipes"

    End If

End Sub

Private Sub cmdVerTrafico_Click()
    frmTrafic.Show

End Sub

Private Sub cmdWorldBackup_Click()

    On Error GoTo ErrHandler

    Me.MousePointer = 11
    FrmStat.Show
    Call ES.DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
    
    Exit Sub

ErrHandler:
    Call LogError("Error en WORLDSAVE")

End Sub

Private Sub cmdRecargarGuardiasPosOrig_Click()

    On Error GoTo ErrHandler

    ReSpawnOrigPosNpcs
    Exit Sub

ErrHandler:
    Call LogError("Error en cmdRecargarGuardiasPosOrig")

End Sub


Private Sub Form_Deactivate()
    frmServidor.Visible = False

End Sub

Private Sub frmAdministracion_Click()
    Me.Visible = False
    frmAdmin.Show

End Sub

Private Sub cmdRecargarAdministradores_Click()
    Call loadAdministrativeUsers

End Sub

Private Sub listDats_Click()
    
    'Chequeamos si hay algun item seleccionado.
    'Lo pongo para prevenir errores.
    If listDats.ListIndex < 0 Then Exit Sub
    
    Select Case UCase$(listDats.Text)
        
        Case "APUESTAS.DAT"
            Call CargaApuestas
            
        Case "ARMASHERRERO.DAT"
            Call LoadArmasHerreria
        
        Case "ARMADURASHERRERO.DAT"
            Call LoadArmadurasHerreria
        
        Case "ARMADURASFACCIONARIAS.DAT"
            Call LoadArmadurasFaccion
        
        Case "BALANCE.DAT"
            Call LoadBalance
        
        Case "CUIDADES.DAT"
            Call CargarCiudades
        
        Case "CONSULTAS.DAT"
            Call ConsultaPopular.LoadData
        
        Case "HECHIZOS.DAT"
            Call CargarHechizos
        
        Case "INVOKAR.DAT"
            Call CargarSpawnList
            
        Case "MOTD.INI"
            Call LoadMotd

        Case "NOMBRESINVALIDOS.txt"
            Call CargarForbidenWords
          
        Case "NPCS.DAT"
            Call CargaNpcsDat(True)

        Case "OBJ.DAT"
            Call LoadOBJData
        
        Case "OBJCARPINTERO.DAT"
            Call LoadObjCarpintero
            
        Case "OBJARTESANO.DAT"
            Call LoadObjArtesano
        
        Case "RETOS.DAT"
            Call LoadArenas
           
        Case "QUESTS.DAT"
            Call LoadQuests
            
        Case Else
            Exit Sub
            
    End Select
    
End Sub
