Attribute VB_Name = "Mobile_FormatExcelIdCost"
Sub Mobile_GetFormatExcelIdCost()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False

    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("O:AC").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select
    
    Range("M2").Select
    Selection.Copy
    
End Sub
