Attribute VB_Name = "Data_PlaySound"
Private Declare Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000

Sub Data_GetPlaySound()
    
    If Sheets("Data").OLEObjects("PlaySoundCheckBox").Object.Value = True Then
        WAVFile = Range("J6").Value
        WAVFile = ThisWorkbook.Path & "\" & WAVFile
        Call PlaySound(WAVFile, 0&, SND_ASYNC Or SND_FILENAME)
    If Sheets("Data").OLEObjects("PlaySoundCheckBox").Object.Value = False Then
        '
    End If
    End If
    
End Sub
