Attribute VB_Name = "Laptops_UnhideAll"
'Unhide all cells

Sub Laptops_GetUnhideAll()
Attribute Laptops_GetUnhideAll.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.CutCopyMode = False
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Range("A1").Select
    Range("F2").Select
    
End Sub

