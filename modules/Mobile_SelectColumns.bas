Attribute VB_Name = "Mobile_SelectColumns"
' This script selects cells depend on how many orders in "Production Order" colomn

Sub Mobile_GetSelectColumns()

    Dim totalOrders As Integer
    totalOrders = Range("C4").Value
    
    ActiveSheet.Range(Cells(2, 5), Cells(totalOrders + 1, 26)).Select
    '2,4 = D2;
    'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")

End Sub

