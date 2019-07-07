Attribute VB_Name = "Mobile_ClearColumns"
' This script clear cells are filled with data

Sub Mobile_GetClearColumns()

Dim totalOrders As Integer
Dim Answer, MyNote As String

totalOrders = Range("C4").Value

    'Message
    MyNote = "Clear Colomns?"
    
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "")
    If Answer = vbYes Then
        'Code for Yes button Press
        
        '2,4 = D2;
        'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")
        
        ActiveSheet.Range(Cells(2, 6), Cells(totalOrders + 2, 9)).Select
        Selection.ClearContents
        
        ActiveSheet.Range(Cells(2, 12), Cells(totalOrders + 2, 16)).Select
        Selection.ClearContents 'remove cells selection
        
        ActiveSheet.Range(Cells(2, 18), Cells(totalOrders + 2, 22)).Select
        Selection.ClearContents
        
        ActiveSheet.Range(Cells(2, 24), Cells(totalOrders + 2, 29)).Select
        Selection.ClearContents
    Else
        'Code for No button Press
    End If
    Range("A1").Select
    Range("F2").Select
    
End Sub

