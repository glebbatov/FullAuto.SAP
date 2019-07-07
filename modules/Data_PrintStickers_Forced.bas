Attribute VB_Name = "Data_PrintStickers_Forced"
Sub Data_GetPrintStickers_Forced()

On Error GoTo Catch
' Get setup with SAP to use the client
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
    
    Dim counter, totalOrders, currentCounter As Integer
    Dim orderNumber As String
    Dim repeat As Long

    'Gets the total number of orders to be processed from Column E row 2
    totalOrders = Range("E2").Value
    counter = 0
    currentCounter = 2

    If totalOrders = 0 Then         'if production orders coloumn does't have any orders
        MsgBox ("No production orders input")
    Else
        'Hit F3(back) button for 5 times to make sure that SAP is on the right page
        For repeat = 1 To 5
        session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
        Next repeat
            'Program looping thru all orders, until the order labels are printed
            Do While counter < totalOrders
                Sheets("Data").Select
                Range("A2").Select
                orderNumber = ActiveCell.Offset(counter, 0).Value
                ActiveSheet.Cells(currentCounter, 1).Select 'highlight current cell
                session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 8  'F8 Print labels
                session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber
                session.findById("wnd[0]").sendVKey 0  'enter
                
                session.findById("wnd[1]/usr/btnBUTTON_1").press  'yes
                
                session.findById("wnd[0]/tbar[1]/btn[9]").press
                session.findById("wnd[0]/usr/btnTC_SERNR_MARK").press
                
                'set E13 to printer name
                printer = Range("E13").Value
                session.findById("wnd[0]/usr/ctxtTSP03D-PADEST").Text = printer
                
                session.findById("wnd[0]/tbar[1]/btn[5]").press
                session.findById("wnd[0]").sendVKey 3  'f3
                session.findById("wnd[0]").sendVKey 3  'f3
                
                counter = counter + 1
                currentCounter = currentCounter + 1 'highlight current cell counter
            Loop
            Range("A2").Select
            MsgBox ("Stickers have been printed!")
Exit Sub

Catch:
619
MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End If

End Sub

