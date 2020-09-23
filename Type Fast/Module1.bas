Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub Bye()
sndPlaySound App.Path & "\sound\bye.wav", &H1
End Sub

Sub Lost()
sndPlaySound App.Path & "\sound\lost.wav", &H1
End Sub

Sub NextWord()
sndPlaySound App.Path & "\sound\next.wav", &H1
End Sub

Sub Start()
sndPlaySound App.Path & "\sound\start.wav", &H1
End Sub

