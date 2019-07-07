Attribute VB_Name = "LaptopsMass_SelectColoumns"
' This script selects cells depend on how many orders in "Production Order" colomn

Sub LaptopsMass_GetSelectColumns()

    Dim totalOrders As Integer
    Sheets("Laptops").Select
    totalOrders = Range("C4").Value
    
    Sheets("Laptops.MassDeployment").Select
    ActiveSheet.Range(Cells(2, 5), Cells(totalOrders + 1, 14)).Select
    '2,4 = D2;
    'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")

End Sub

