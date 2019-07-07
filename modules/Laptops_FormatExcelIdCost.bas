Attribute VB_Name = "Laptops_FormatExcelIdCost"
'Hides coloumns, ready to pull ID/CostCenter#

Sub Laptops_GetFormatExcelIdCost()

    Application.CutCopyMode = False
    
    'unhide all coloumns before script
    Cells.Select
    Selection.EntireColumn.Hidden = False

    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:N").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:R").Select
    Selection.EntireColumn.Hidden = True
    Columns("T:T").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select

    Range("O2").Select
    Selection.Copy
    
End Sub
