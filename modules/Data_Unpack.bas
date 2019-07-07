Attribute VB_Name = "Data_Unpack"
' This script will "unpack" orders using quantity number (number of items in order)
' previously pulled using QuantityCheck script

Sub Data_GetUnpack()

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
    
    
    Dim counter, totalOrders, orderQuantity, currentCounter As Integer
    Dim orderNumber, quantity As String
    Dim repeat As Long
        
    Sheets("Data").Select
    orderQuantity = Range("B2").Value
        
            If orderQuantity <= 0 Then
                'if production orders coloumn doesn't have any orders
                MsgBox "No quantity input." & vbNewLine & "Please pull order quantity first"
                
            Else
        
                'Gets the total number of orders to be processed from Column E row 2
                totalOrders = Range("E2").Value
                counter = 0
                currentCounter = 2
            
                If totalOrders = 0 Then         'if production orders coloumn does't have any orders
                    MsgBox ("No production orders input")
                Else
                    'Message
                    'MyNote = "Ready to Unpack?"
                    
                    'if user choose "YES" in message box,
                    'program looping thru all the orders,
                    'until all the orders are "packed"
                    
                    'Display MessageBox
                    'Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "")
                    'If Answer = vbYes Then
                    
                        'Hit F3(back) button for 5 times to make sure that SAP is on the right page
                        For repeat = 1 To 5
                        session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                        Next repeat
                            'Program looping thru all orders, until all orders get "unpack"
                            Do While counter < totalOrders
                                Sheets("Data").Select
                                Range("A2").Select
                                orderNumber = ActiveCell.Offset(counter, 0).Value
                                ActiveSheet.Cells(currentCounter, 1).Select 'highlight current cell
                                session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
                                session.findById("wnd[0]").sendVKey 0   'Enter
                                session.findById("wnd[0]/tbar[1]/btn[2]").press     'F2 Pack
                                
                                session.findById("wnd[0]/usr/ctxtGV_AUFNR").Text = orderNumber  'Prod.Order field
                                Application.Wait (Now + TimeValue("00:00:01")) '1 second delay
                                session.findById("wnd[0]/usr/btn%#AUTOTEXT006").press   'F5:Start
                                
                                session.findById("wnd[0]/usr/ctxtGV_AUFNR").Text = orderNumber    'Prod.Order field
                                Application.Wait (Now + TimeValue("00:00:01")) '1 second delay
                                session.findById("wnd[0]/usr/btn%#AUTOTEXT008").press   'FinnishButton
                                
                                Range("B2").Select
                                orderQuantity = ActiveCell.Offset(counter, 0).Value
                                ActiveSheet.Cells(currentCounter, 2).Select  'highlight current cell
                                session.findById("wnd[0]/usr/txtGV_MGVRG").Text = orderQuantity
                                session.findById("wnd[0]/usr/btnFINUPDATE").press   'Fin.UpdateButon
                                
                                session.findById("wnd[0]/tbar[0]/btn[3]").press
                                session.findById("wnd[0]/tbar[0]/btn[3]").press
                                counter = counter + 1
                                currentCounter = currentCounter + 1 'highlight current cell counter
                            Loop
                            Range("A2").Select
                            'MsgBox ("Unpacked!")
                            'Else
                        'Code for No button Press
                    End If
                Exit Sub
        Sheets("Data").Select
        Range("A2").Select
    
        End If
'End If
    
Catch:
619
MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End Sub

