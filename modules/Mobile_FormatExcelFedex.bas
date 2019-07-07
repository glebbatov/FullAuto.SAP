Attribute VB_Name = "Mobile_FormatExcelFedex"
'Hides coloumns, ready to print FedEx lables

Sub Mobile_GetFormatExcelFedex()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False

    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:M").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:X").Select
    Selection.EntireColumn.Hidden = True
    Columns("Z:AC").Select
    Selection.EntireColumn.Hidden = True
    
    Range("AK1").Select
    
    Range("J2:K2").Select
    
    ActiveWindow.SmallScroll ToRight:=5
    
End Sub


