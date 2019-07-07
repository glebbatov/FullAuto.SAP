Attribute VB_Name = "Laptops_FormatExcelScan"
'Hides coloumns, ready to data scan

Sub Laptops_GetFormatExcelScan()

    Dim totalOrders As Integer
    totalOrders = Range("C4").Value

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False

    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:M").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:R").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
    
    'Range("A1").Select
    Range("S2").Select
    'ActiveWindow.SmallScroll ToRight:=2
    
    ActiveSheet.Range(Cells(2, 7), Cells(totalOrders + 1, 14)).Select
    '2,4 = D2;
    'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")
    
End Sub
