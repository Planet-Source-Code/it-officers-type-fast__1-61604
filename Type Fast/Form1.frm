VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Type Fast"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   4200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2760
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000012&
      Caption         =   "Hard"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "Medium"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "Easy"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Play Game"
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   3360
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Make Your Choice :  "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "YOUR SCORE :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "TYPE THE WORD HERE: "
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5350
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   0
      X2              =   6360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "wassim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   2640
      Picture         =   "Form1.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This Program is written by ' TEA MAKER '
' Feel free to use the code as you like
' Or Email me : Tea_makers@hotmail.com

Dim word As Integer
Dim i, j, k  As Integer
Dim v As Integer

Private Sub Command1_Click()
Module1.Start
NextWord
If Option1.Value = True Then
v = 1
End If
If Option2.Value = True Then
v = 2
End If
If Option3.Value = True Then
v = 3
End If

Command1.Visible = False
Timer4.Enabled = True
Label5.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
End Sub

Private Sub Form_Load()
Option1.Value = True
i = 120
j = 840
End Sub

Private Sub Form_Unload(Cancel As Integer)
Module1.Bye
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Text1.Text = Label1.Caption And Timer1.Enabled = True Then
Text1.Text = ""
Timer1.Enabled = False
i = 120
j = 840
Label4.Caption = Label4.Caption + 1
Module1.NextWord
NextWord
Timer1.Enabled = True
End If

If Text1.Text = Label1.Caption And Timer2.Enabled = True Then
Text1.Text = ""
Timer2.Enabled = False
i = 120
j = 840
Label4.Caption = Label4.Caption + 2
Module1.NextWord
NextWord
Timer2.Enabled = True
End If

If Text1.Text = Label1.Caption And Timer3.Enabled = True Then
Text1.Text = ""
Timer3.Enabled = False
i = 120
j = 840
Label4.Caption = Label4.Caption + 3
Module1.NextWord
NextWord
Timer3.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
i = i + 15
j = j + 15
Image1.Top = i
Label1.Top = j
If Label1.Top >= 4560 Then
Module1.Lost
MsgBox "You lost! You Have Scored " & Label4.Caption & " Points", vbExclamation, "Hard Luck"
Text1.Text = ""
Timer1.Enabled = False
Command1.Visible = True

Label5.Visible = True
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Line1.Visible = False
Text1.Visible = False
Image1.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
i = i + 25
j = j + 25
Image1.Top = i
Label1.Top = j
If Label1.Top >= 4560 Then
Module1.Lost
MsgBox "You lost! You Have Scored " & Label4.Caption & " Points", vbExclamation, "Hard Luck"
Text1.Text = ""
Timer2.Enabled = False

Command1.Visible = True
Label5.Visible = True
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Line1.Visible = False
Text1.Visible = False
Image1.Visible = False
End If
End Sub

Private Sub Timer3_Timer()
i = i + 35
j = j + 35
Image1.Top = i
Label1.Top = j
If Label1.Top >= 4560 Then
Module1.Lost
MsgBox "You lost! You Have Scored " & Label4.Caption & " Points", vbExclamation, "Hard Luck"
Text1.Text = ""
Timer3.Enabled = False

Command1.Visible = True
Label5.Visible = True
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Line1.Visible = False
Text1.Visible = False
Image1.Visible = False
End If

End Sub

Private Sub Timer4_Timer()
Label6.Visible = True
k = k + 1
If k = 1 Then
Label6.Caption = "Ready"
End If
If k = 2 Then
Label6.Caption = "Set"
End If
If k = 3 Then
Label6.Caption = "Go"
End If
If k = 4 Then
Label6.Visible = False


If v = 1 Then
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False

End If
If v = 2 Then
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = False

End If

If v = 3 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True

End If
' Finish
Label5.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Line1.Visible = True
Text1.Visible = True
Image1.Visible = True

Label4.Caption = 0
i = 120
j = 840
k = 0
Timer4.Enabled = False
End If
End Sub

Private Function NextWord()
word = Int(Rnd * 50)
If word = 0 Then
Label1.Caption = "beautiful"
End If
If word = 1 Then
Label1.Caption = "something"
End If
If word = 2 Then
Label1.Caption = "Washington"
End If
If word = 3 Then
Label1.Caption = "twilight"
End If
If word = 4 Then
Label1.Caption = "Beirut"
End If
If word = 5 Then
Label1.Caption = "love"
End If
If word = 6 Then
Label1.Caption = "television"
End If
If word = 7 Then
Label1.Caption = "microsoft"
End If
If word = 8 Then
Label1.Caption = "karim"
End If
If word = 9 Then
Label1.Caption = "visual basic"
End If
If word = 10 Then
Label1.Caption = "kiss me"
End If
If word = 11 Then
Label1.Caption = "don't mess"
End If
If word = 12 Then
Label1.Caption = "intel"
End If
If word = 13 Then
Label1.Caption = "direct x"
End If
If word = 14 Then
Label1.Caption = "Italy"
End If
If word = 15 Then
Label1.Caption = "marketing"
End If
If word = 16 Then
Label1.Caption = "computer"
End If
If word = 17 Then
Label1.Caption = "worms"
End If
If word = 18 Then
Label1.Caption = "internet"
End If
If word = 19 Then
Label1.Caption = "racing"
End If
If word = 20 Then
Label1.Caption = "children"
End If
If word = 21 Then
Label1.Caption = "anything"
End If
If word = 22 Then
Label1.Caption = "butterfly"
End If
If word = 23 Then
Label1.Caption = "scream"
End If
If word = 24 Then
Label1.Caption = "feeling"
End If
If word = 25 Then
Label1.Caption = "identical"
End If
If word = 26 Then
Label1.Caption = "immortal"
End If
If word = 27 Then
Label1.Caption = "eating"
End If
If word = 28 Then
Label1.Caption = "budget"
End If
If word = 29 Then
Label1.Caption = "Mark"
End If
If word = 30 Then
Label1.Caption = "bullet"
End If
If word = 31 Then
Label1.Caption = "bottle"
End If
If word = 32 Then
Label1.Caption = "essay"
End If
If word = 33 Then
Label1.Caption = "cannot"
End If
If word = 34 Then
Label1.Caption = "angels"
End If
If word = 35 Then
Label1.Caption = "cartoon"
End If
If word = 36 Then
Label1.Caption = "masking"
End If
If word = 37 Then
Label1.Caption = "drunk"
End If
If word = 38 Then
Label1.Caption = "traveling"
End If
If word = 39 Then
Label1.Caption = "Lebanon"
End If
If word = 40 Then
Label1.Caption = "freedom"
End If
If word = 41 Then
Label1.Caption = "Britain"
End If
If word = 42 Then
Label1.Caption = "package"
End If
If word = 43 Then
Label1.Caption = "outlandish"
End If
If word = 44 Then
Label1.Caption = "football"
End If
If word = 45 Then
Label1.Caption = "java script"
End If
If word = 46 Then
Label1.Caption = "flamingo"
End If
If word = 47 Then
Label1.Caption = "potatoes"
End If
If word = 48 Then
Label1.Caption = "practice"
End If
If word = 49 Then
Label1.Caption = "tv tuner"
End If
If word = 50 Then
Label1.Caption = "mercury"
End If

End Function
