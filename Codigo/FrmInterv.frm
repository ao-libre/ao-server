VERSION 5.00
Begin VB.Form FrmInterv 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Suerte 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Suerte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   55
      Top             =   4320
      Width           =   2055
      Begin VB.TextBox txtMinarSuerte 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   58
         Text            =   "6"
         Top             =   720
         Width           =   450
      End
      Begin VB.TextBox txtTalarSuerte 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   57
         Text            =   "6"
         Top             =   720
         Width           =   450
      End
      Begin VB.TextBox txtPescarSuerte 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Text            =   "6"
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mineria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   61
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Talar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   60
         Top             =   360
         Width           =   495
      End
      Begin VB.Label pesca 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pesca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Guardar Intervalos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NPCs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      TabIndex        =   49
      Top             =   2160
      Width           =   1695
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A.I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   150
         TabIndex        =   50
         Top             =   240
         Width           =   1365
         Begin VB.TextBox txtAI 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   52
            Text            =   "0"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.TextBox txtNPCPuedeAtacar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   135
            TabIndex        =   51
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   54
            Top             =   840
            Width           =   195
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puede atacar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   255
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clima && Ambiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      TabIndex        =   39
      Top             =   2160
      Width           =   2865
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Frio y Fx Ambientales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   2625
         Begin VB.TextBox txtCmdExec 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   44
            Text            =   "0"
            Top             =   1110
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloPerdidaStaminaLluvia 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1320
            TabIndex        =   43
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloWAVFX 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   42
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloFrio 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   41
            Text            =   "0"
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TimerExec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   48
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stamina Lluvia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   47
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FxS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   270
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   195
            TabIndex        =   45
            Top             =   840
            Width           =   345
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   90
         TabIndex        =   24
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloParaConexion 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   26
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtTrabajo 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   25
            Text            =   "0"
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IntervaloCon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trabajo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   660
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Combate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   1545
         TabIndex        =   19
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtPuedeAtacar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   22
            Text            =   "0"
            Top             =   1200
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloLanzaHechizo 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   20
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puede Atacar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   930
            Width           =   1170
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lanza Spell"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   285
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hambre y sed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5925
         TabIndex        =   14
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloHambre 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloSed 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hambre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   255
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   17
            Top             =   930
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sanar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   4470
         TabIndex        =   9
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtSanaIntervaloDescansar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Text            =   "0"
            Top             =   480
            Width           =   1050
         End
         Begin VB.TextBox txtSanaIntervaloSinDescansar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descansando"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sin descansar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   12
            Top             =   930
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Stamina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   3015
         TabIndex        =   4
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtStaminaIntervaloSinDescansar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   6
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtStaminaIntervaloDescansar 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   165
            TabIndex        =   5
            Text            =   "0"
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sin descansar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descansando"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   255
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Duracion Spells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   2400
         Begin VB.TextBox txtInvocacion 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1170
            TabIndex        =   37
            Text            =   "0"
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloInvisible 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1170
            TabIndex        =   34
            Text            =   "0"
            Top             =   495
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloParalizado 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   31
            Text            =   "0"
            Top             =   1170
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloVeneno 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   30
            Text            =   "0"
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invocacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   38
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invisible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   35
            Top             =   285
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paralizado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Veneno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   300
            Width           =   660
         End
      End
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir (Esc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
End
Attribute VB_Name = "FrmInterv"
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

Public Sub AplicarIntervalos()

    '?????????? Intervalos del main loop ????????
    SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
    StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
    SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
    StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
    IntervaloSed = val(txtIntervaloSed.Text)
    IntervaloHambre = val(txtIntervaloHambre.Text)
    IntervaloVeneno = val(txtIntervaloVeneno.Text)
    IntervaloParalizado = val(txtIntervaloParalizado.Text)
    IntervaloInvisible = val(txtIntervaloInvisible.Text)
    IntervaloFrio = val(txtIntervaloFrio.Text)
    IntervaloWavFx = val(txtIntervaloWAVFX.Text)
    IntervaloInvocacion = val(txtInvocacion.Text)
    IntervaloParaConexion = val(txtIntervaloParaConexion.Text)

    '///////////////// TIMERS \\\\\\\\\\\\\\\\\\\

    IntervaloUserPuedeCastear = val(txtIntervaloLanzaHechizo.Text)
    IntervaloUserPuedeTrabajar = val(txtTrabajo.Text)
    IntervaloUserPuedeAtacar = val(txtPuedeAtacar.Text)
    
    '///////////////////Suerte \\\\\\\\\\\\\\\\\\
    
    DificultadPescar = val(txtPescarSuerte.Text)
    DificultadTalar = val(txtTalarSuerte.Text)
    DificultadMinar = val(txtMinarSuerte.Text)
End Sub

Private Sub Command1_Click()

    On Error Resume Next

    Call AplicarIntervalos

End Sub

Private Sub Command2_Click()

    On Error GoTo Err

    'Intervalos
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", str(SanaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", str(StaminaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", str(SanaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", str(StaminaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", str(IntervaloSed))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", str(IntervaloHambre))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", str(IntervaloVeneno))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", str(IntervaloParalizado))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", str(IntervaloInvisible))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", str(IntervaloFrio))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", str(IntervaloWavFx))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", str(IntervaloInvocacion))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", str(IntervaloParaConexion))

    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", str(IntervaloUserPuedeCastear))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo", str(IntervaloUserPuedeTrabajar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", str(IntervaloUserPuedeAtacar))
    
    '&&&&&&&&&&&&&&&&&&&& Suerte desde INIT &&&&&&&&&&&&&&&&&&&&&&&&&
    
    Call WriteVar(IniPath & "Server.ini", "DIFICULTAD", "DificultadPescar", str(DificultadPescar))
    Call WriteVar(IniPath & "Server.ini", "DIFICULTAD", "DificultadTalar", str(DificultadTalar))
    Call WriteVar(IniPath & "Server.ini", "DIFICULTAD", "DificultadMinar", str(DificultadMinar))
    
    
    MsgBox "Los intervalos se han guardado sin problemas."

    Exit Sub
Err:
    MsgBox "Error al intentar grabar los intervalos"

End Sub

Private Sub Form_Load()
    With Me
        .txtPescarSuerte.Text = DificultadPescar
        .txtTalarSuerte.Text = DificultadTalar
        .txtMinarSuerte.Text = DificultadMinar
        .txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
        .txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
        .txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
        .txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
        .txtIntervaloSed.Text = IntervaloSed
        .txtIntervaloHambre.Text = IntervaloHambre
        .txtIntervaloVeneno.Text = IntervaloVeneno
        .txtIntervaloParalizado.Text = IntervaloParalizado
        .txtIntervaloInvisible.Text = IntervaloInvisible
        .txtIntervaloFrio.Text = IntervaloFrio
        .txtIntervaloWAVFX.Text = IntervaloWavFx
        .txtInvocacion.Text = IntervaloInvocacion
        .txtIntervaloParaConexion.Text = IntervaloParaConexion
        .txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
        .txtTrabajo.Text = IntervaloUserPuedeTrabajar
        .txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
    End With
    
End Sub

Private Sub ok_Click()
    Me.Visible = False

End Sub

