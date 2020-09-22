VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN Messenger Assistant Setup"
   ClientHeight    =   5340
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Advanced"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check for software updates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2640
      TabIndex        =   13
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4695
      Begin VB.TextBox AwayMess 
         Height          =   735
         Left            =   240
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "Setup.frx":0000
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox checkAnswer 
         Caption         =   "Answering Machine"
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
         TabIndex        =   8
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Outgoing message"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      Begin VB.CheckBox CheckExcited 
         Caption         =   "Excited speech!!!"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox checkWelcome 
         Caption         =   "Welcome speech"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox checkSpeech 
         Caption         =   "Speech"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox checkCharm 
         Caption         =   "Hourly charm speech"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox checkIM 
         Caption         =   "Incomming IM speech"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox checkStatus 
         Caption         =   "User status speech"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "makes computer sound excited"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "greeting from your computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "tells you the time every hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "incomming im text reading"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "online/offline/away notification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   630
      Left            =   120
      Picture         =   "Setup.frx":0036
      ScaleHeight     =   570
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AwayMess_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Away Message", AwayMess
Form1.AwayMess = AwayMess.text
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\VersionCheck", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\VersionCheck", "0"
End If

End Sub

Private Sub AwayMess_GotFocus()
AwayMess.SelStart = 0: AwayMess.SelLength = Len(AwayMess.text)
End Sub

Private Sub checkAnswer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If checkAnswer.Value = 1 Then
Form1.Check1.Value = 1
Form1.Check1_Click
Else
Form1.Check1.Value = 0
Form1.Check1_Click
End If
End Sub

Private Sub checkCharm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If checkCharm.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Charm", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Charm", "0"
End If
End Sub


Private Sub CheckEgnor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CheckEgnor.Value = 1 Then CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\EgnorOn", "1" Else CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\EgnorOn", "0"

End Sub

Private Sub CheckExcited_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If CheckExcited.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Excited", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Excited", "0"
End If
End Sub

Private Sub checkIM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If checkIM.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\IM", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\IM", "0"
End If
End Sub

Private Sub checkSpeech_Click()
If checkSpeech.Value = 1 Then
Form1.Check2.Value = 1
Form1.Check2_Click
Else
Form1.Check2.Value = 0
Form1.Check2_Click
End If
End Sub

Private Sub checkStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If checkStatus.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status", "0"
End If


End Sub

Private Sub checkWelcome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


If checkWelcome.Value = 1 Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Welcome", "1"
Else
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Welcome", "0"
End If
End Sub




Private Sub Command1_Click()
Form4.Show
Unload Form3
End Sub

Private Sub Command3_Click()
Form3.Visible = False

If Check1.Value = 1 Then
If Internet.IsConnected = True Then
Dim Temp50 As String
Dim CurrentVersion As String
CurrentVersion = "2.17 with improved speech settings"

Temp50 = IsNewVersion(CurrentVersion, Form1.Inet1, "http://www.wilsonr1.karoo.net/msnaver.txt")
If Not Temp50 = vbNullString Then
response$ = MsgBox("MSN Assistant version " & Temp50 & " is now available!" & vbCrLf & "Would you like to upgrade to this version?", vbQuestion + vbYesNo, "Program Update")
If response$ = vbYes Then Internet.DownloadFile "http://www.wilsonr1.karoo.net/MSNAssistant.zip"
End If
End If
End If



Unload Form3
End Sub

Private Sub Form_Load()
Dim temp2 As String


temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Away Message")
If temp2 = "" Then GoTo 10
AwayMess = temp2
10

temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status")
If temp2 = "0" Then checkStatus.Value = 0 Else checkStatus.Value = 1

temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\IM")
If temp2 = "0" Then checkIM.Value = 0 Else checkIM.Value = 1

temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Charm")
If temp2 = "0" Then checkCharm.Value = 0 Else checkCharm.Value = 1

temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Welcome")
If temp2 = "0" Then checkWelcome.Value = 0 Else checkWelcome.Value = 1

temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Excited")
If temp2 = "1" Then CheckExcited.Value = 1 Else CheckExcited.Value = 0

Dim temp1 As String
temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Speech")
If temp1 = "1" Then checkSpeech.Value = 1: Form2.SpeechON.Checked = True
If temp1 = "0" Then checkSpeech.Value = 0: Form2.SpeechON.Checked = False
If temp1 = "" Then checkSpeech.Value = 1: Form2.SpeechON.Checked = True
If Form1.Check1.Value = 1 Then checkAnswer.Value = 1


temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\VersionCheck")
If Not temp2 = "0" Then Check1.Value = 1 Else Check1.Value = 0

Form3.Visible = True


End Sub

