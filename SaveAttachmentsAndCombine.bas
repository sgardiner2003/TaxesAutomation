Sub SaveAttachmentsAndCombine()
    Dim olItem As Object
    Dim olAttachment As Attachment
    Dim saveFolder As String
    Dim logFilePath As String
    Dim fileName As String
    Dim filePath As String
    Dim i As Integer
    Dim n As Integer
    Dim SNAM As String
    Dim SNAMlist As String
    Dim fso As Object
    Dim logFile As Object
    Dim folder As Object
    Dim file As Object
    Dim xlApp As Object
    Dim wb As Object
    Dim ws1, TickersSold As Object
    Dim masterwb As Object
    Dim r As Integer
    Dim masterpath As String
    
    On Error GoTo CleanUp

    ''' Paths '''
    saveFolder = "C:\Users\sgardiner\TEMP STUFF" ' ADD PATH TO YOUR STUFF FOLDER HERE!
    taxesFolder = saveFolder & "\taxes - backup"
    logFilePath = saveFolder & "\SNAMs.txt"

    ' Check if they chose an existing save folder
    If Dir(saveFolder, vbDirectory) = "" Then
        MsgBox "Folder does not exist: " & saveFolder
        Exit Sub
    End If
    
    ' Check if taxes folder exists already
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(taxesFolder) Then
        On Error Resume Next
        fso.DeleteFolder taxesFolder, True
        On Error GoTo CleanUp
    End If
    fso.CreateFolder taxesFolder
    
    ' Initialize master_sheet and xl application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    Set masterwb = xlApp.Workbooks.Add
    
    masterpath = saveFolder & "\master_sheet.xlsx"
    
    ' Delete master_sheet.xlsx if it exists already and save a new one
    If Dir(masterpath) <> "" Then
        On Error Resume Next
        Kill masterpath
        On Error GoTo CleanUp
    End If
    
    On Error Resume Next
    masterwb.SaveAs masterpath, FileFormat:=51 ' format 51 is xlsx
    If Err.Number <> 0 Then
        MsgBox "Please close previous master sheet or tax sheets!"
        Err.Clear
        GoTo CleanUp
    End If
    On Error GoTo CleanUp

    ''' Pulling tax attachments '''

    ' SNAMlist will be printed to a text file at the end
    SNAMlist = ""
    n = 1

    For Each olItem In Application.ActiveExplorer.Selection
        If olItem.Attachments.Count > 0 Then
            SNAM = Right(Trim(olItem.Subject), 6)
            SNAM = Replace(SNAM, "\", "-") ' clean up just in case
            SNAMlist = SNAMlist & SNAM & vbCrLf

            For Each olAttachment In olItem.Attachments
                fileName = olAttachment.fileName
                fileName = SNAM & Mid(fileName, InStrRev(fileName, "."))
                filePath = taxesFolder & "\" & fileName

                On Error Resume Next
                olAttachment.SaveAsFile filePath
                If Err.Number <> 0 Then
                    MsgBox "Please close previous master sheet or tax sheets!"
                    Err.Clear
                    GoTo CleanUp
                End If
            Next
        End If
        n = n + 1
    Next

    ' Write log file
    Set logFile = fso.CreateTextFile(logFilePath, True)
    logFile.Write SNAMlist
    logFile.Close
    
    ''' Merge sheets into one workbook'''
    
    Set folder = fso.GetFolder(taxesFolder)
    
    ' Iterate over each file in taxes folder, copy info and paste into master_sheet
    r = 1
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
            Set wb = xlApp.Workbooks.Open(file.Path)
            Set ws1 = wb.Sheets(1) ' SNAM sheet
            Set TickersSold = wb.Sheets(2) ' TickersSold sheet
            
            ' Initialize blank sheets
            If r = 1 Then
                masterwb.Worksheets.Add After:=masterwb.Worksheets(1)
            Else
                masterwb.Worksheets.Add After:=masterwb.Worksheets(2 * r - 2)
                masterwb.Worksheets.Add After:=masterwb.Worksheets(2 * r - 1)
            End If
            
            ' Copy info from sheet to master sheet
            ws1.Cells.Copy Destination:=masterwb.Worksheets(2 * r - 1).Range("A1")
            TickersSold.Cells.Copy Destination:=masterwb.Worksheets(2 * r).Range("A1")
            
            ' Rename sheets with SNAM
            masterwb.Worksheets(2 * r - 1).Name = ws1.Name
            masterwb.Worksheets(2 * r).Name = ws1.Name & " tickers sold"
            
            ' Open Immediate window to see these
            Debug.Print "Processing file: " & file.Name
            Debug.Print "Master sheet count: " & masterwb.Worksheets.Count
            Debug.Print "Current SNAM: " & ws1.Name
            
            ' Close wb every time to prevent being unable to delete files
            wb.Close SaveChanges:=False
            r = r + 1
        End If
    Next file
    
    ''' Closing out '''
    masterwb.Close SaveChanges:=True
    Set masterwb = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "Attachments and SNAMs saved to: " & logFilePath
    Exit Sub
    
    ' For when workbooks are left open:
CleanUp:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    
    On Error Resume Next
    If Not masterwb Is Nothing Then
        masterwb.Close SaveChanges:=True
        Set masterwb = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    
End Sub
