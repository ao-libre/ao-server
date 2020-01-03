VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form frmCargando 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3180
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267.49
   ScaleMode       =   0  'User
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar cargar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   1
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " aa"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmCargando"
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

Private Sub Form_Load()

    Label1(2).Caption = GetVersionOfTheServer()
    
    Picture1.Picture = LoadPicture(App.Path & "\logo.jpg")
    
    Call Me.VerifyIfUsingLastVersion
    
End Sub

Function VerifyIfUsingLastVersion()

    On Error Resume Next
    
    Dim AOUpdater_Path As String
        AOUpdater_Path = App.Path & "\Autoupdate.exe"

    If FileExist(AOUpdater_Path, vbNormal) Then
    
        If Not (CheckIfRunningLastVersion) Then
            
            If MsgBox("Tu version no es la actual, Deseas ejecutar el actualizador automatico?.", vbYesNo) = vbYes Then
                Call ShellExecute(Me.hWnd, "open", AOUpdater_Path, vbNullString, vbNullString, 1)
                End
            End If
    
        End If
        
    End If
    
End Function

Private Function CheckIfRunningLastVersion() As Boolean

    On Error GoTo errorHandler

    'Declaramos los objetos a usar.
    Dim JsonObject     As Object
    Dim Inet           As clsInet: Set Inet = New clsInet
    
    Dim responseGithub As String
    Dim versionNumberMaster As String, versionNumberLocal As String
    
    'Nos conectamos a GitHub
    responseGithub = Inet.OpenRequest("https://api.github.com/repos/ao-libre/ao-server/releases/latest", "GET")
    responseGithub = Inet.Execute
    responseGithub = Inet.GetResponseAsString
    
    'Chequeamos si recibimos algo en primer lugar.
    If LenB(responseGithub) <> 0 Then
        
        'Trato de parsear el JSON obtenido a traves del control Inet.
        Set JsonObject = JSON.parse(responseGithub)
        
        'Si hay algun error, devolvemos FALSE.
        If LenB(JSON.GetParserErrors) <> 0 Then GoTo errorHandler
        
        'Comparamos la version obtenida de GitHub con la local.
        versionNumberMaster = JsonObject.Item("tag_name")
        versionNumberLocal = GetVar(App.Path & "\Server.ini", "INIT", "VersionTagRelease")
        
        If versionNumberMaster = versionNumberLocal Then
            Set JsonObject = Nothing
            Set Inet = Nothing
            
            CheckIfRunningLastVersion = True
        End If
        
    End If
    
    'Si llegamos a este punto significa que algo paso.
    GoTo errorHandler

errorHandler:

    'Liberamos los recursos.
    Set JsonObject = Nothing
    Set Inet = Nothing
    
    If Err.Number Then
        Call LogError("Error: " & Err.Number & " Descripcion: " & Err.description)
    End If
    
    CheckIfRunningLastVersion = False

End Function
