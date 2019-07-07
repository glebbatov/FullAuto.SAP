Attribute VB_Name = "Mobile_PullData"
' This script is pulling different data from "zint" into several columns:
' Sales Order, User Name, E-Mail, and Shipping Address
' using "Production Order" number

Sub Mobile_GetPullData()

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
    
        Dim counter, totalOrders, pos As Integer
        Dim orderNumber, Temp, UserEmail As String
        Dim repeat As Long
        
        totalOrders = Range("C4").Value
        counter = 0
        
        'press protection message
        MyNote = "Pull Data?"
        'Display MessageBox
        Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "")
        If Answer = vbYes Then
            If totalOrders = 0 Then         'if production orders coloumn does't have any orders
                MsgBox ("No production orders input")
            Else
                'Hit F3(back) button for 5 times to make sure that SAP is on the right page
                For repeat = 1 To 5
                session.findById("wnd[0]/tbar[0]/btn[3]").press     'F3 back
                Next repeat
    
                    Do While counter < totalOrders
                    
                        Sheets("Mobiles").Select
                        Range("F2").Select  'production order
                        orderNumber = ActiveCell.Offset(counter, 0).Value
                        
                        session.findById("wnd[0]/tbar[0]/okcd").Text = "zint"
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/tbar[1]/btn[9]").press
                        session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").Text = orderNumber
                        session.findById("wnd[0]/usr/ctxtAFKO-AUFNR").caretPosition = 10
                        session.findById("wnd[0]").sendVKey 0
                        
                        Range("G2").Select 'sales order
                        ActiveCell.Offset(counter, 0) = session.findById("wnd[0]/usr/txtV_KDAUF").Text
                        
                        Range("L2").Select  'used id
                        Temp02 = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp02 Like "*User ID:*" Then         'only go for "User ID" if field is there
                            Temp02 = Replace(Temp02, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp02 = Replace(Temp02, "                                                                                                                                                                                                                                         .", "")
                            Temp02 = Replace(Temp02, "                                                                                                                                                                                         ", "")
                            Temp02 = Replace(Temp02, "                                             .", "")
                            Temp02 = Replace(Temp02, "~", " ")
                            TolVal = Len(Temp02)
                            pos2 = InStr(Temp02, "User ID:")
                                If Temp02 Like "*Cost Center:*" Then
                                    pos1 = InStr(Temp02, "Cost Center:")
                                Else 'Temp02 Like "*PO and Line #:*" Then
                                    pos1 = InStr(Temp02, "PO and Line #:")
                                End If
                            PosLenth = pos1 - pos2
                            UserId = Mid(Temp02, pos2, PosLenth)
                            UserId = Application.WorksheetFunction.Clean(UserId)
                            UserId = LTrim(UserId)
                            RevUserID = (Len(UserId)) - 8
                            UserId = Right(UserId, RevUserID)
                            ActiveCell.Offset(counter, 0) = Trim(UserId)
                        End If
                        
                        Range("N2").Select  'CostCenter#
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*Cost Center:*" Then         'only go for "Cost Center:" if field is there
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                            TolVal = Len(Temp)
                            pos1 = InStr(Temp, "PO and Line #:")
                            pos2 = InStr(Temp, "Cost Center:")
                            PosLenth = pos1 - pos2
                            UserName = Mid(Temp, pos2, PosLenth)
                            UserName = Application.WorksheetFunction.Clean(UserName)
                            UserName = LTrim(UserName)
                            RevUserName = (Len(UserName)) - 12
                            UserName = Right(UserName, RevUserName)
                            ActiveCell.Offset(counter, 0) = Trim(UserName)
                         End If
                        
                        Range("AC2").Select  'user name
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*User Name:*" Then         'only go for "User Name" if field is there
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                            TolVal = Len(Temp)
                            pos1 = InStr(Temp, "User Email:")
                            pos2 = InStr(Temp, "User Name:")
                            PosLenth = pos1 - pos2
                            UserName = Mid(Temp, pos2, PosLenth)
                            UserName = Application.WorksheetFunction.Clean(UserName)
                            UserName = LTrim(UserName)
                            RevUserName = (Len(UserName)) - 10
                            UserName = Right(UserName, RevUserName)
                            ActiveCell.Offset(counter, 0) = Trim(UserName)
                        End If
                        
                        Range("M2").Select  'email
                        Temp1 = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp1 Like "*User Email:*" Then         'only go for "User Email" if field is there
                            LeftVal = InStr(Temp, "@")
                            RightVal = TolVal - LeftVal
                            LeftMail = Left(Temp, LeftVal - 1)
                            LeftMailx = InStrRev(Left(Temp, LeftVal - 1), ":")
                            NameEmailVal = LeftVal - LeftMailx
                            UserEmail = Right(LeftMail, NameEmailVal - 1)
                            UserEmail = Application.WorksheetFunction.Clean(UserEmail)
                            UserEmail = LTrim(UserEmail)
                            RightMail = Right(Temp, RightVal)
                            RightTolVal = Len(RightMail)
                            RightMailx = InStr(Left(RightMail, RightTolVal), " ")
                            EmailAddress = Left(RightMail, RightMailx - 1)
                            ActiveCell.Offset(counter, 0) = UserEmail & "@" & EmailAddress
                        End If
                            
                        Range("Y2").Select 'shipping address
                        session.findById("wnd[0]/usr/txtV_AUFNR").SetFocus
                        session.findById("wnd[0]").sendVKey 2
                        session.findById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtAFPOD-KDAUF").SetFocus
                        session.findById("wnd[0]").sendVKey 2
                        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").SetFocus
                        session.findById("wnd[0]").sendVKey 2
                        counter2 = 0
                        rowName = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & counter2 & "]").Text
                        
TryAgain:
                        If Not rowName = "Ship-to party" Then
                            counter2 = counter2 + 1
                            rowName = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & counter2 & "]").Text
                            GoTo TryAgain
                        End If
                        
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-STREET[3," & counter2 & "]").SetFocus
                        session.findById("wnd[0]").sendVKey 2
                        streetName = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").Text
                        
                        City = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").Text
                        State = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").Text
                        ZipCode = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").Text
                        CostCenterAddress = streetName & " / " & City & " , " & State & " " & ZipCode
                        
                        ActiveCell.Offset(counter, 0) = CostCenterAddress
                        session.findById("wnd[1]").Close
                        
                        session.findById("wnd[0]").sendVKey 3
                        session.findById("wnd[0]").sendVKey 3
                        session.findById("wnd[0]").sendVKey 3
                
                        
                        Range("O2").Select  'area code
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*Area Code:*" Then         'only go for "area code" if field is there
                            'replace text and spaces to ""(nothing) from SAP textBox
                            'Replace(string, "replace", "replacer" )
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                    
                            TolVal = Len(Temp)            'Len() - return a number of characters in a string
                            pos1 = InStr(Temp, "User Name:")
                            pos2 = InStr(Temp, "Area Code:")
                            PosLenth = pos1 - pos2
                            areacode = Mid(Temp, pos2, PosLenth)
                            areacode = Application.WorksheetFunction.Clean(areacode)
                            areacode = LTrim(areacode)
                            Revareacode = (Len(areacode)) - 10
                            areacode = Right(areacode, Revareacode)
                            ActiveCell.Offset(counter, 0) = Trim(areacode)
                        End If
                        
                        
                        Range("P2").Select  'carrier
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*Carrier:*" Then         'only go for "carrier" if field is there
                            'replace text and spaces to ""(nothing) from SAP textBox
                            'Replace(string, "replace", "replacer" )
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                    
                            TolVal = Len(Temp)            'Len() - return a number of characters in a string
                            
                                If Temp Like "*Phone number:*" Then         'if "phone number" follows "carrier" field
                                    pos1 = InStr(Temp, "Phone number:")
                                    pos2 = InStr(Temp, "Carrier:")
                                    PosLenth = pos1 - pos2
                                    carrier = Mid(Temp, pos2, PosLenth)
                                    carrier = Application.WorksheetFunction.Clean(carrier)
                                    carrier = LTrim(carrier)
                                    Revcarrier = (Len(carrier)) - 8
                                    carrier = Right(carrier, Revcarrier)
                                    ActiveCell.Offset(counter, 0) = Trim(carrier)
                                Else                                      'if "area code" follows "carrier" field
                                    pos1 = InStr(Temp, "Area Code:")
                                    pos2 = InStr(Temp, "Carrier:")
                                    PosLenth = pos1 - pos2
                                    carrier = Mid(Temp, pos2, PosLenth)
                                    carrier = Application.WorksheetFunction.Clean(carrier)
                                    carrier = LTrim(carrier)
                                    Revcarrier = (Len(carrier)) - 8
                                    carrier = Right(carrier, Revcarrier)
                                    ActiveCell.Offset(counter, 0) = Trim(carrier)
                                End If
                        End If
                
                
                        Range("S2").Select  'mobile number
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*Phone number:*" Then         'only go for "phone number" if field is there
                            'replace text and spaces to ""(nothing) from SAP textBox
                            'Replace(string, "replace", "replacer" )
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                    
                            TolVal = Len(Temp)            'Len() - return a number of characters in a string
                            pos1 = InStr(Temp, "User Name:")
                            pos2 = InStr(Temp, "Phone number:")
                            PosLenth = pos1 - pos2
                            mobile = Mid(Temp, pos2, PosLenth)
                            mobile = Application.WorksheetFunction.Clean(mobile)
                            mobile = LTrim(mobile)
                            RevMobile = (Len(mobile)) - 13
                            mobile = Right(mobile, RevMobile)
                            ActiveCell.Offset(counter, 0) = Trim(mobile)
                        End If
                        
                        
                        Range("V2").Select  'PO# & Line Item
                        Temp = session.findById("wnd[0]/usr/cntlLABTEXT/shellcont/shell").Text
                        If Temp Like "*PO and Line*" Then         'only go for "PO and Line #" if field is there
                            'replace text and spaces to ""(nothing) from SAP textBox
                            'Replace(string, "replace", "replacer" )
                            Temp = Replace(Temp, "Text on Sales Order Header:                                                                                                                                                                                                                                    ", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                                                                         .", "")
                            Temp = Replace(Temp, "                                                                                                                                                                                         ", "")
                            Temp = Replace(Temp, "                                             .", "")
                            Temp = Replace(Temp, "~", " ")
                            TolVal = Len(Temp)            'Len() - return a number of characters in a string
                            pos1 = InStr(Temp, "***")
                            pos2 = InStr(Temp, "PO and Line #:")
                            PosLenth = pos1 - pos2
                            PO = Mid(Temp, pos2, PosLenth)
                            PO = Application.WorksheetFunction.Clean(PO)
                            PO = LTrim(PO)
                            RevPO = (Len(PO)) - 14
                            PO = Right(PO, RevPO)
                            ActiveCell.Offset(counter, 0) = Trim(PO)
                        End If
                        counter = counter + 1
                        session.findById("wnd[0]").sendVKey 3
                        session.findById("wnd[0]").sendVKey 3
                        session.findById("wnd[0]").sendVKey 3
                    Loop
                Range("A1").Select  'make spreadsheet look nice again
                Range("F2").Select  'make spreadsheet look nice
                MsgBox ("Process Complete")
            End If
    Else
            'Code for No button Press
    End If
    
End Sub

