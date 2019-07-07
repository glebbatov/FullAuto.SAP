Attribute VB_Name = "Laptops_SAPSortColumns"
' This script sorts colomns in "enter data" subpage

Sub Laptops_GetSAPSortColumns()

On Error GoTo Catch

    If Not IsObject(sapApplication) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set sapApplication = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = sapApplication.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject sapApplication, "on"
    End If
    
    Dim repeat As Long
    
    Range("F2").Select
    orderNumber = ActiveCell.Offset(counter, 0).Value
        
    
    
    'Hit F3(back) button for 10 times to make sure that SAP is on the right page
    For repeat = 1 To 5
    session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
    Next repeat
        session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[1]/btn[9]").press
        session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_SERNR/txtWA_SERNR-SERNR[1,0]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").reorderTable "0 3 5 6 7 1 2 4 8 9 10 11 12 13"
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(5).Width = 11
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(6).Width = 8
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(1).Width = 5
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(2).Width = 7
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(3).Width = 8
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR").Columns.elementAt(4).Width = 8
        session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
        'return back and come back again for correct field change while scanning
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_SERNR/txtWA_SERNR-SERNR[1,0]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        session.findById("wnd[0]/usr/tblZWMM_LABPROCESSTC_COMP_SERNR/txtWA_COMP_SERIAL-BOX[1,0]").SetFocus
        'highlight "box number" field
    Exit Sub
    
Catch:
619

MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End Sub
