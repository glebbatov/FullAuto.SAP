Attribute VB_Name = "Data_ClearColomns"
' This script clear cells that filled

Sub Data_GetClearColumns()

Dim totalOrders As Integer
Dim Answer, MyNote As String

totalOrders = Range("E2").Value

    'Message
    MyNote = "Clear Colomns?"
    
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "")
    If Answer = vbYes Then
        'Code for Yes button Press
        '2,4 = D2;
        'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")
        ActiveSheet.Range(Cells(2, 1), Cells(totalOrders + 2, 2)).Select
        Selection.ClearContents
    Else
        'Code for No button Press
    End If
    Range("A2").Select

End Sub

