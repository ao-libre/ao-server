VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administración del servidor"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Echar todos los PJS no privilegiados"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "R"
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cboPjs 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Echar"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmAdmin.frm
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

Option Explicit

Private Sub cboPjs_Change()
Call ActualizaPjInfo
End Sub

Private Sub cboPjs_Click()
Call ActualizaPjInfo
End Sub

Private Sub Command1_Click()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(tIndex).name & " ha sido echado.", FontTypeNames.FONTTYPE_SERVER))
    Call CloseSocket(tIndex)
End If

End Sub

Public Sub ActualizaListaPjs()
Dim LoopC As Long

With cboPjs
    .Clear
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                .AddItem UserList(LoopC).name
                .ItemData(.NewIndex) = LoopC
            End If
        End If
    Next LoopC
End With

End Sub

Private Sub Command3_Click()
Call EcharPjsNoPrivilegiados

End Sub

Private Sub ActualizaPjInfo()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
    With UserList(tIndex)
        Text1.Text = .outgoingData.length & " elementos en cola." & vbCrLf
    End With
End If

End Sub
