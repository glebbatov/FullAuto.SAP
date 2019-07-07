Attribute VB_Name = "Mobile_FormatExcelScan"
'Hides coloumns, ready to data scan

Sub Mobile_GetFormatExcelScan()
Attribute Mobile_GetFormatExcelScan.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim totalOrders As Integer
    totalOrders = Range("C4").Value

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:T").Select
    Selection.EntireColumn.Hidden = True
    Range("V:AA").Select
    Selection.EntireColumn.Hidden = True
    Range("AC:AC").Select
    Selection.EntireColumn.Hidden = True
    
    'Range("A1").Select
    Range("AD2").Select
    
    ActiveSheet.Range(Cells(2, 8), Cells(totalOrders + 1, 29)).Select
    '2,4 = D2;
    'totalOrders+1, 16 = P + totalOrders+1 (increment starts from 1, so "totalOrders+1")
    
    ActiveWindow.SmallScroll ToRight:=0
    
End Sub
