Attribute VB_Name = "media"
'----------------------------------------------------------
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1         '  play asynchronously

Public Sub tocarSom()
  Call sndPlaySound(App.path & "\som.wav", SND_ASYNC)
End Sub
