Attribute VB_Name = "vbPlayWav"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal ipszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_Sync = &H0
Public Const SND_Async = &H1
Public Const SND_NoDefault = &H2
Public Const SND_Loop = &H8
Public Const SND_NoStop = &H10
Public Sub PlayWav(ByVal SoundName As String)
    Dim ret As Variant
    Flag% = SND_Async Or SND_NoDefault
    ret = sndPlaySound(SoundName, Flag%)
    DoEvents
End Sub
Public Sub StopWav()
    Dim ret As Variant
    ret = sndPlaySound(0&, 0&)
End Sub
