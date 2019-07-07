Attribute VB_Name = "Laptops_SAPFormatExcelEnterData"
'Hides coloumns, ready to data enter

Sub Laptops_GetSAPFormatExcelEnterData()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:L").Select
    Selection.EntireColumn.Hidden = True
    Columns("N:R").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select
    
    Range("I2:M2").Select
    Selection.Copy
    
End Sub

