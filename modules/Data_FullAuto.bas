Attribute VB_Name = "Data_FullAuto"
Sub Data_GetFullAuto()
    
    'Plays audio file
    If Sheets("Data").OLEObjects("PlaySoundFullAuto").Object.Value = True Then Call Data_PlaySound.Data_GetPlaySound
    
    Call Data_PrintLabNotes.Data_GetPrintLabNotes
    Call Data_PullOrderQuantity.Data_GetPullOrderQuantity
    Call Data_Unpack.Data_GetUnpack
    Call Data_PrintStickers.Data_GetPrintStickers
    
    MsgBox ("Expressed!")
    
End Sub
