VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False 'disable screen update to save time
Application.Calculation = xlCalculationManual 'disable calcs to save time

    Dim vendorlist As Worksheet
        
    Dim fNameAndPath As Variant
    Dim oldworkbook As Workbook
    Dim reusewrksht As Worksheet
    
    Dim workingbook As Workbook
    
    Dim checkNameAndPath As Variant
    Dim checklist As Workbook
    Dim checksheet As Worksheet
    
    Set workingbook = ThisWorkbook
    
    'Ask for, and open, the last quarterly report
    MsgBox ("Please select the analysis for this category for last quarter.")
    fNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the last quarter's analysis?")
    If fNameAndPath = False Then
        Set reusewrksht = Nothing
        Set oldworkbook = Nothing
    Else
        Set oldworkbook = Workbooks.Open(fNameAndPath, True, True)
        Set reusewrksht = oldworkbook.Sheets("All Data")
    End If
    
    'Ask for, and open, the list of checks to use
    MsgBox ("Please select the list of recent checks.")
    checkNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the list of recent checks?")
    If checkNameAndPath = False Then
        Set checksheet = Nothing
        Set checklist = Nothing
    Else
        Set checklist = Workbooks.Open(checkNameAndPath, True, True)
        Set checksheet = checklist.Sheets(1)
    End If

    'Check to see if user forgot to insert vendor column and insert it if it's missing
    If workingbook.Worksheets("Paste Data Here").Range("N1").Value = "Control2" Then
        workingbook.Worksheets("Paste Data Here").Range("N1").EntireColumn.Insert
        workingbook.Worksheets("Paste Data Here").Range("N1").Value = "Vendor Name"
    End If
    
    Set vendorlist = workingbook.Worksheets("Vendor List")
    
    'Create Vendor Dictionary
    numvenrows = vendorlist.UsedRange.Rows.Count
    Dim vendict As New Scripting.Dictionary
    
    For j = 2 To numvenrows 'loop through all the vendors in the vendor list and
            If (InStr(1, "Do Not Use", vendorlist.Cells(j, 2).Value, 1) > 0) Then
                vendict.Add Key:=j, Item:=vendorlist.Cells(j, 23).Value
            Else
                vendict.Add Key:=j, Item:=vendorlist.Cells(j, 2).Value
            End If
    Next j
    
    Dim totaldeb As Double
    Dim matchdeb As Double
    Dim totalcred As Double
    Dim matchcred As Double
    Dim lineval As Double
    
    'Get the row count and start the timer for the match
    numrows = workingbook.Worksheets("Paste Data Here").UsedRange.Rows.Count
    starttime = Timer
    
    'Convert all text to numbers if appropriate
    workingbook.Worksheets("Paste Data Here").Columns(13).NumberFormat = "0"
    workingbook.Worksheets("Paste Data Here").Columns(13).Value = workingbook.Worksheets("Paste Data Here").Columns(13).Value
    workingbook.Worksheets("Paste Data Here").Columns(15).NumberFormat = "0"
    workingbook.Worksheets("Paste Data Here").Columns(15).Value = workingbook.Worksheets("Paste Data Here").Columns(15).Value
    workingbook.Worksheets("Paste Data Here").Columns(18).NumberFormat = "0"
    workingbook.Worksheets("Paste Data Here").Columns(18).Value = workingbook.Worksheets("Paste Data Here").Columns(18).Value
    
    'Run the vendor match on all rows
    For i = 2 To numrows
        If (i Mod 500 = 0) Then DoEvents
        lineval = workingbook.Worksheets("Paste Data Here").Cells(i, 12).Value
        If (lineval > 0) Then totaldeb = totaldeb + lineval
        If (lineval < 0) Then totalcred = totalcred + lineval
        If IsEmpty(workingbook.Worksheets("Paste Data Here").Cells(i, 14)) Then
            workingbook.Worksheets("Paste Data Here").Cells(i, 14).Value = _
                VendorMatch(workingbook.Worksheets("Paste Data Here").Rows(i), vendict, vendorlist, reusewrksht, checksheet)
            If Not (workingbook.Worksheets("Paste Data Here").Cells(i, 14).Value = "") Then
                If (lineval > 0) Then matchdeb = matchdeb + lineval
                If (lineval < 0) Then matchcred = matchcred + lineval
            End If
        End If
        If (i Mod 500 = 0) Then Application.StatusBar = "Updating " & CStr(Int((i / numrows) * 100)) & "%" 'Show the progress for the match in the status bar in %
    Next i
    Application.Calculation = xlCalculationAutomatic 'turn back on the calculation
    finaltime = Int(Timer - starttime) 'figure out how long the match took
    
    Application.ScreenUpdating = True 'turn back on screen updating
    
    'Figure out the total debits and credits
    'Figure out the total debits and credits that were matched
    'Calculate the percentages
    
    ratedeb = Int((matchdeb / totaldeb) * 100)
    ratecred = Int((matchcred / totalcred) * 100)
    
    
    nummatches = WorksheetFunction.CountA(workingbook.Worksheets("Paste Data Here").Range("N:N")) - 1 'get the number of matches that were made
    
    'close the workbooks for last quarter and checks
    If Not oldworkbook Is Nothing Then oldworkbook.Close (False)
    If Not checklist Is Nothing Then checklist.Close (False)
    
    Set oldworkbook = Nothing
    Set checklist = Nothing
    Set vendorlist = Nothing
    
    'Display a message box to let user know that match is finished and show the stats
    'MsgBox ("Finished! " & vbNewLine & nummatches & " matched. That's a " & matchrate & "% success rate." & vbNewLine & "The program ran for " & finaltime & " seconds.")
    MsgBox ("Finished! " & vbNewLine & nummatches & " matched. We matched " & ratedeb & "% of the expenses and " & ratecred & "% of the credits." & vbNewLine & "The program ran for " & finaltime & " seconds.")
    Application.StatusBar = False 'fix the status bar
    
End Sub

Private Sub CommandButton2_Click()
    Worksheets("Paste Data Here").Cells.Clear 'Clear the worksheet for the next match

End Sub



