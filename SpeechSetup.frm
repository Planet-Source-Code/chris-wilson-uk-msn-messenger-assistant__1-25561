VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN Messenger Assistant Speech Setup"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4995
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customize Speech Events"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   4695
      Begin VB.TextBox txtIM 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "(%1) says"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtWelcome 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "(%3) (%4)"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtIdle 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Text            =   "(%1) is now idle"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtOffline 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Text            =   "(%1) has gone offline"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtOnline 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "(%1) has come online"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtAvailable 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Text            =   "(%1) is now available"
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txtLunch 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   "(%1) has gone out to lunch"
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txtPhone 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "(%1) is now on the phone"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtBRB 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Text            =   "(%1) will be right back"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtBusy 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Text            =   "(%1) is now busy"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtAway 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Text            =   "(%1) has gone away"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label17 
         Caption         =   "Incoming IM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Welcome speech:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Idle:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Go offline"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Come online"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Available:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Out to lunch:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "On the phone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Be right back:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Busy:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Away:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      Picture         =   "SpeechSetup.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   4695
      TabIndex        =   12
      Top             =   120
      Width           =   4755
   End
   Begin VB.Label Label13 
      Caption         =   "(%4) = Local friendly name"
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "(%3) = Good Morn/Noon/Evening"
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Use the following commands for speech"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "(%2) = Users hotmail ID"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "(%1) = Users friendly name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Advanced Speech Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   4575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form4
End Sub

Private Sub Form_Load()
txtWelcome = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtWelcome")
txtOnline = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOnline")
txtOffline = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOffline")
txtAway = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAway")
txtBRB = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBRB")
txtPhone = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtPhone")
txtAvailable = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAvailable")
txtBusy = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBusy")
txtIdle = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIdle")
txtLunch = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtLunch")
txtIM = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIM")

End Sub

Private Sub txtAvailable_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAvailable", txtAvailable
End Sub

Private Sub txtAway_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAway", txtAway
End Sub

Private Sub txtBRB_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBRB", txtBRB
End Sub

Private Sub txtBusy_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBusy", txtBusy
End Sub

Private Sub txtIdle_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIdle", txtIdle
End Sub

Private Sub txtIM_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIM", txtIM
End Sub

Private Sub txtLunch_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtLunch", txtLunch
End Sub

Private Sub txtOffline_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOffline", txtOffline
End Sub

Private Sub txtOnline_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOnline", txtOnline
End Sub

Private Sub txtPhone_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtPhone", txtPhone
End Sub

Private Sub txtWelcome_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtWelcome", txtWelcome

End Sub
