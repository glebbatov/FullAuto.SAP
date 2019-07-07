Attribute VB_Name = "Laptops_FormatExcelFedex"
'Hides coloumns, ready to print FedEx lables

Sub Laptops_GetFormatExcelFedex()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select
    
    Range("J2").Select
    Selection.Copy
    
    ActiveWindow.SmallScroll ToRight:=4
End Sub

