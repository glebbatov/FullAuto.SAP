Attribute VB_Name = "Laptops_CopyCurrentPos"
' This script pulls "Production order" numbers from "Data" sheet

Sub Laptops_GetCopyCurrentPos()
Attribute Laptops_GetCopyCurrentPos.VB_ProcData.VB_Invoke_Func = " \n14"

    'select/copy A colomn the way down till last filled cell, start with cell A2
    Sheets("Data").Select
    totalOrders = Range("E2").Value
    
    If totalOrders = 0 Then         'if production orders coloumn doesn't have any orders
        MsgBox ("No production orders input")
    Else
        Range("A2", Range("A2").End(xlDown)).Select
        Selection.Copy
        
        Sheets("Laptops").Select
        PO = Range("C4").Value
        If PO > 0 Then
            Answer = MsgBox("Replace current POs?", vbQuestion + vbYesNo, "")
            If Answer = vbYes Then
                    'paste copied cells to E colonm, start with cell E2
                    Sheets("Laptops").Select
                    Range("F2").Select
                    ActiveSheet.Paste
                    
                    'make all sheets cell selection looks nice
                    Sheets("Data").Select
                    Application.CutCopyMode = False
                    Range("A2").Select
                    Sheets("Laptops").Select
                    Range("F2").Select
                End If
        Else
            Sheets("Laptops").Select
            Range("F2").Select
            ActiveSheet.Paste
            
            'make all sheets cell selection looks nice
            Sheets("Data").Select
            Application.CutCopyMode = False
            Range("A2").Select
            Sheets("Laptops").Select
            Range("F2").Select
        End If
    End If
    
    Sheets("Laptops").Select
    Range("F2").Select
    
End Sub
