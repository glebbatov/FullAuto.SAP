Attribute VB_Name = "Data_PullCNFCheck"
' This script will take a production order number,
' copy it to "enter data" field, and stays on
' "enter data" page as long as the value of E18 cell(seconds) is.
' It gives time to check if order was CNFed and also check its status

Sub Data_GetPullCnfCheck()

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
    
    Dim counter, totalOrders, delay, seconds, currentCounter As Integer
    Dim orderNumber, CNF As String
    Dim repeat As Long

    'Gets the total number of orders to be processed from E2 cell
    totalOrders = Range("E2").Value
    counter = 0
    currentCounter = 2

    If totalOrders = 0 Then         'if production orders coloumn does't have any orders
        MsgBox ("No production orders input")
    Else
    
    'Plays audio file
    If Sheets("Data").OLEObjects("PlaySoundCNF").Object.Value = True Then Call Data_PlaySound.Data_GetPlaySound
    
        'Hit F3(back) button for 5 times to make sure that SAP is on the right page
        For repeat = 1 To 5
        session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
        Next repeat
        
            Do While counter < totalOrders
                Sheets("Data").Select
                
                Range("A2").Select
                orderNumber = ActiveCell.Offset(counter, 0).Value
                ActiveSheet.Cells(currentCounter, 1).Select 'highlight current cell
                session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[1]/btn[9]").press
                session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber
                session.findById("wnd[0]").sendVKey 0
                
                delay = TimeSerial(0, 0, Range("E19").Value)    'delay value (sec) assigned to cell E19
                Application.Wait (Now + delay)  'delay time when "enter data" page appear
                
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                
                counter = counter + 1
                currentCounter = currentCounter + 1 'highlight current cell counter
            Loop
            
            Range("A2").Select
            MsgBox ("Here you go!")
            
Exit Sub

Catch:
619
MsgBox "Stopped." & vbNewLine & "Please, set SAP to the default page"

End If

End Sub

