Attribute VB_Name = "Mobile_SAPFormatExcelEnterData"
'Hides coloumns, ready to data enter

Sub Mobile_GetSAPFormatExcelEnterData()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("M:AC").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select
    
    Range("L2").Select
    Selection.Copy
    
End Sub


