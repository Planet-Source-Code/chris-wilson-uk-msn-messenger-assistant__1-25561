VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "MSN Menu Form"
   ClientHeight    =   480
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   2760
   Icon            =   "MSN Menu.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   480
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MSNMENU 
      Caption         =   "&MSNMenu"
      Begin VB.Menu Show 
         Caption         =   "&Show"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu Setup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu ViewLog 
         Caption         =   "&View Log"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu AnswerON 
         Caption         =   "&Answer on"
      End
      Begin VB.Menu SpeechON 
         Caption         =   "&Speech on"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu Cancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AnswerON_Click()

If Form1.Check1.Value = 0 Then
Form1.Check1.Value = 1
AnswerON.Checked = True
Else
Form1.Check1.Value = 0
AnswerON.Checked = False
End If

Form1.Check1_Click

End Sub

Private Sub ChangeOGM_Click()
Dim temp1 As String
temp1 = InputBox("Enter your new outgoing message:", "MSN Answering Machine", Form1.AwayMess.Text)
If temp1 = "" Then Exit Sub
Form1.AwayMess = temp1
End Sub

Private Sub Exit_Click()
Form1.ExitProgram
End Sub

Private Sub Setup_Click()
Form3.Show
End Sub

Private Sub Show_Click()
Form1.Show: Form1.SetFocus
End Sub

Private Sub SpeechON_Click()
If Form1.Check2.Value = 0 Then
Form1.Check2.Value = 1
SpeechON.Checked = True
Else
Form1.Check2.Value = 0
SpeechON.Checked = False
End If

Form1.Check2_Click

End Sub


Private Sub ViewLog_Click()
Shell "notepad.exe c:\windows\msn_messenger.log", vbNormalFocus

End Sub
