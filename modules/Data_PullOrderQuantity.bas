Attribute VB_Name = "Data_PullOrderQuantity"
' This script will pull order quantity from SAP using the "zint" page ("Enter Data" subpage)
' This number will be used further while unpack and pack scripts

Sub Data_GetPullOrderQuantity()

On Error GoTo Catch

    'Get setup with SAP to use the client
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
    Dim orderNumber, quantity As String
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
        'Program looping thru all orders, locate "Order Quantity" field, and copy it to B2 colonm
        Do While counter < totalOrders
            
            Sheets("Data").Select
            
            Range("A2").Select
            orderNumber = ActiveCell.Offset(counter, 0).Value
            ActiveSheet.Cells(currentCounter, 1).Select  'highlight current cell
            session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
            session.findById("wnd[0]").sendVKey 0   'enter
            session.findById("wnd[0]/tbar[1]/btn[9]").press 'F9 Enter Data
            session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber    'paste current order number
            session.findById("wnd[0]").sendVKey 0   'enter
            
            Range("B2").Select
            session.findById("wnd[0]/usr/txtV_GAMNG").SetFocus  'highlight "Order Quantity" field
            quantity = session.findById("wnd[0]/usr/txtV_GAMNG").Text 'assign varQuantity to "Order Quantity" value
            ActiveCell.Offset(counter, 0) = quantity 'copy varQuantity value to an active cell (B colomn)
            ActiveSheet.Cells(currentCounter, 2).Select  'highlight current cell
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 Back
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            counter = counter + 1
            currentCounter = currentCounter + 1 'highlight current cell counter
        Loop
        
        Range("A2").Select
        'MsgBox ("Here you go!")
        
    Exit Sub
    
Catch:
619
MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End If

End Sub
