Function SelectAPLSession(selectedName As String) As Object
    Dim Sys As Object
    Dim Sess As Object
    Dim i As Integer
    Dim sessionList As String
    Dim found As Boolean

    Set Sys = CreateObject("Extra.System")

    If Sys.Sessions.Count = 0 Then
        MsgBox "No EXTRA! sessions are currently open."
        Exit Function
    End If

    found = False
    For i = 0 To Sys.Sessions.Count - 1
        If Sys.Sessions.Item(i).Name = selectedName Then
            Set Sess = Sys.Sessions.Item(i)
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "Session '" & selectedName & "' not found."
        Exit Function
    End If

    Set SelectAPLSession = Sess
End Function

Sub Audit()
    Dim Sess As Object
    Dim lastRow As Long
    Dim snam As String
    Dim cusips As Variant
    Dim totalLines As Integer
    Dim screenText As String
    Dim lineText As Variant
    Dim found As Boolean
    Dim page As Integer
    Dim i As Integer
    
    ' --- Initialize session ---
    Set Sess = SelectAPLSession("Lord Abbett")
    If Sess Is Nothing Then Exit Sub
    
    ' --- Inputs ---
    
    lastRow = Worksheets("Sheet1").Cells(Rows.Count, "B").End(xlUp).Row
    Debug.Print (lastRow)
    
    Dim rng As Range
    
    snam = CStr(Worksheets("Sheet1").Range("A3").Value)
    
    Set rng = Worksheets("Sheet1").Range("B3:B" & lastRow)
    If rng.Cells.Count = 1 Then
        ReDim cusips(1 To 1, 1 To 1)
        cusips(1, 1) = rng.Value
    Else
        cusips = rng.Value
    End If

    ReDim names(LBound(cusips, 1) To UBound(cusips, 1))
    Debug.Print "Rows in cusips: " & UBound(cusips, 1)
    Debug.Print "Cols in cusips: " & UBound(cusips, 2)

    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "AUDIT<Enter>"
    Sess.Screen.WaitHostQuiet 1000
    
    If InStr(Sess.Screen.GetString(23, 1, 40), "AS OF DATE") = 0 Then
        MsgBox "Screen not ready for AUDIT. Press escape."
        Exit Sub
    End If
    
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys snam & "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    
    If InStr(Sess.Screen.GetString(22, 1, 10), "0 RECORDS") > 0 Then
        MsgBox "SNAM not found: " & snam
        Exit Sub
    End If
    
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "N<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "BROWSE<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    
    totalLines = 24
    
    For i = LBound(cusips, 1) To UBound(cusips, 1)
        found = False
        page = 1
        
        upAttempts = 0
        
        Do
            If InStr(Sess.Screen.GetString(1, 1, 20), "LORD, ABBETT & CO.") > 0 Then
                Exit Do
            Else
                Sess.Screen.SendKeys "<up>"
                Sess.Screen.WaitHostQuiet 500
                upAttempts = upAttempts + 1
                If upAttempts > 5 Then
                    MsgBox "Could not find top of report."
                    Exit Do
                End If
            End If
        Loop
        
        Do
            Sess.Screen.WaitHostQuiet 500
            
            screenText = Sess.Screen.GetString(1, 1, totalLines * 80)
            
            If InStr(screenText, CStr(cusips(i, 1))) > 0 Then
                For j = 0 To totalLines - 1
                    lineText = Mid(screenText, j * 80 + 1, 80)
                    If InStr(lineText, cusips(i, 1)) > 0 Then
                        found = True
                        securityName = Trim(Mid(lineText, 12, 29))
                        Worksheets("Sheet1").Range("C" & i + 2).Value = securityName
                        Exit For
                    End If
                Next j
                Exit Do
            End If
            
            If InStr(screenText, "GRAND TOTAL") = 0 Then
                Sess.Screen.SendKeys "<down>"
                Sess.Screen.WaitHostQuiet 500
                page = page + 1
            Else
                Exit Do
            End If
        Loop
            
        If Not found Then
            MsgBox "CUSIP: " & cusips(i, 1) & " not found for SNAM: " & snam
            If MsgBox("Stop macro?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                Exit Sub
            End If
        End If
        
        Debug.Print (securityName)
        
        Next i

End Sub

Sub BWTX()
    Dim Sess As Object
    Dim lastRow As Long
    Dim snam As String
    Dim names As Variant
    Dim cost As String
    Dim val As String
    Dim origCost As Double
    Dim markVal As Double
    Dim total As Double
    Dim screenText As String
    Dim lineText As Variant
    Dim nameRow As Integer
    Dim found As Boolean
    Dim page As Integer
    Dim i As Integer
    Dim upAttempts As Integer
    Dim GLtotal As Double
    
        ' --- Initialize session ---
    Set Sess = SelectAPLSession("Lord Abbett")
    If Sess Is Nothing Then Exit Sub
    
    ' --- Inputs ---
    
    lastRow = Worksheets("Sheet1").Cells(Rows.Count, "C").End(xlUp).Row
    
    Dim rng As Range
    
    snam = CStr(Worksheets("Sheet1").Range("A3").Value)
    
    Set rng = Worksheets("Sheet1").Range("C3:C" & lastRow)
    If rng.Cells.Count = 1 Then
        ReDim names(1 To 1, 1 To 1)
        names(1, 1) = rng.Value
    Else
        names = rng.Value
    End If
    
    ' Create report
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "BWTX<Enter>"
    Sess.Screen.WaitHostQuiet 500
    
    If InStr(Sess.Screen.GetString(23, 1, 33), "PRESENT REPORT IN WHICH CURRENCY?") = 0 Then
        MsgBox "Screen not ready for BWTX."
        Exit Sub
    End If
    
    Sess.Screen.SendKeys "USD<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys snam & "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    
    If InStr(Sess.Screen.GetString(23, 1, 18), "FUNCTION COMPLETED") > 0 Then
        MsgBox "SNAM not found: " & snam
        Exit Sub
    End If
    
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "N<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "BROWSE<Enter>"
    Sess.Screen.WaitHostQuiet 500
    Sess.Screen.SendKeys "<Enter>"
    Sess.Screen.WaitHostQuiet 500
    
    ' Scrape screen
    
    GLtotal = 0
    
    For i = LBound(names) To UBound(names)
        found = False
        total = 0
        ' GLtotal is the total G/L for every security/cusip, total is just for one security/cusip
        secname = UCase(CStr(names(i, 1)))
        
        upAttempts = 0
        
        Do
            If InStr(Sess.Screen.GetString(1, 1, 20), "LORD, ABBETT & CO.") > 0 Then
                Exit Do
            Else
                Sess.Screen.SendKeys "<up>"
                Sess.Screen.WaitHostQuiet 200
                Sess.Screen.SendKeys "<left>"
                Sess.Screen.WaitHostQuiet 200
                upAttempts = upAttempts + 1
                If upAttempts > 10 Then
                    MsgBox "Could not find LORD, ABBETT & CO. after scrolling up."
                    Exit Do
                End If
            End If
        Loop
        
        Do
            screenText = Sess.Screen.GetString(1, 1, 1920)
            
            If InStr(screenText, secname) > 0 Then
                Dim totalLines As Integer
                totalLines = Len(screenText) \ 80
                
                For j = 0 To totalLines - 1
                    lineText = Mid(screenText, j * 80 + 1, 80)
                    If InStr(lineText, secname) > 0 Then
                        found = True
                        nameRow = j + 1
                        Sess.Screen.SendKeys "<right>"
                        Sess.Screen.WaitHostQuiet 200
                        cost = Sess.Screen.GetString(nameRow, 26, 10)
                        val = Sess.Screen.GetString(nameRow, 53, 8)
                        Debug.Print (cost)
                        Debug.Print (val)
                        origCost = CDbl(Replace(Trim(cost), ",", ""))
                        markVal = CDbl(Replace(Trim(val), ",", ""))
                        total = markVal - origCost
                        GLtotal = GLtotal + total
                        Worksheets("Sheet1").Range("D" & i + 2).Value = total
                        Sess.Screen.SendKeys "<left>"
                        Sess.Screen.WaitHostQuiet 200
                        Exit For
                    End If
                Next j
                Exit Do
            End If
            
            If InStr(screenText, "GRAND TOTAL") = 0 Then
                Sess.Screen.SendKeys "<down>"
                Sess.Screen.WaitHostQuiet 200
            Else
                Exit Do
            End If
        Loop
            
        If Not found Then
            MsgBox "CUSIP: " & secname & " not found for SNAM: " & snam
            If MsgBox("Stop macro?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                Exit Sub
            End If
        End If
        
        Next i
        
        Worksheets("Sheet1").Range("E3").Value = GLtotal
        
End Sub

Sub Reset()
    Columns("A:E").ClearContents
    Range("A2").Value = "SNAMs"
    Range("B2").Value = "CUSIPs"
    Range("C2").Value = "Sec. Names"
    Range("D2").Value = "Indiv. Gain/Loss"
    Range("E2").Value = "Total Gain/Loss"
    Range("A3").Select
End Sub


