Attribute VB_Name = "Mobile_CopyCurrentPos"
' This script pulls "Production order" numbers from "Data" sheet

Sub Mobile_GetCopyCurrentPos()

    'select/copy A colomn the way down till last filled cell, start with cell A2
    Sheets("Data").Select
    totalOrders = Range("E2").Value
    
    If totalOrders = 0 Then         'if production orders coloumn doesn't have any orders
        MsgBox ("No production orders input")
    Else
        Range("A2", Range("A2").End(xlDown)).Select
        Selection.Copy
        
        Sheets("Mobiles").Select
        PO = Range("C4").Value
        If PO > 0 Then
            Answer = MsgBox("Replace current POs?", vbQuestion + vbYesNo, "")
            If Answer = vbYes Then
                'paste copied cells to E colonm, start with cell E2
                Sheets("Mobiles").Select
                Range("F2").Select
                ActiveSheet.Paste
    
                'make all sheets cell selection looks nice
                Sheets("Data").Select
                Application.CutCopyMode = False
                Range("A2").Select
                Sheets("Mobiles").Select
                Range("F2").Select
                End If
        Else
            Sheets("Mobiles").Select
            Range("F2").Select
            ActiveSheet.Paste
            
            'make all sheets cell selection looks nice
            Sheets("Data").Select
            Application.CutCopyMode = False
            Range("A2").Select
            Sheets("Mobiles").Select
            Range("F2").Select
        End If
    End If

    Sheets("Data").Select
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("Mobiles").Select
    Range("F2").Select

End Sub

