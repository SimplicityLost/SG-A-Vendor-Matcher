Attribute VB_Name = "a_VendorMatcher"
Option Explicit

Function VendorMatch(inputrow, vendict As Scripting.Dictionary, vendorlist As Worksheet, Optional oldwrksht As Variant = False, Optional checklist As Worksheet = Nothing) As String
    Dim j As Integer
    Dim doc_desc As String
    Dim det_desc As String
    Dim control1 As String
    Dim control2 As String
    Dim refnum As String
    Dim acct_desc As String
    Dim acct_date As Date
          
    Dim oldven As String
    Dim venname As String
    
    
    'Grab important sections of Row and assign them to variables
    For j = 1 To 26
        If IsError(inputrow(j)) Then VendorMatch = "": Exit Function
    Next j
    doc_desc = inputrow(9)
    det_desc = inputrow(18)
    control1 = inputrow(13)
    control2 = inputrow(15)
    refnum = inputrow(8)
    acct_desc = inputrow(11)
    acct_date = inputrow(4)
    
    
    'Check to see if we've already assigned a vendor last quarter and pass it up.
    If Not oldwrksht = False Then
        oldven = reuseoldven(refnum, acct_date, inputrow(1), inputrow(12), oldwrksht)
        If Not (oldven = "") Then VendorMatch = oldven: Exit Function
    End If
     
    'Categorize each entry based on relevant features
    Select Case True
    
    'Specific Vendor Cases that are common enough but will take longer to do in more rigorous searches (to save time)
    
    Case (InStr(1, doc_desc, "ACCR", 1) > 1)
        VendorMatch = "ACCRUAL"
        
    Case (InStr(1, doc_desc, "ecova", 1) > 0 Or InStr(1, refnum, "ecova", 1) > 0)
        VendorMatch = "ECOVA INC"
        
    Case (det_desc = "WEBSITE")
        VendorMatch = "FACTORY WEBSITE FEES"
    
    Case (InStr(1, control1, "photon", 1) > 0)
        VendorMatch = "PHOTON CONCEPTS"
    
    Case (InStr(1, control1, "TRUE", 1) > 0)
        VendorMatch = "TRUE CAR INC"
        
    Case (InStr(1, control2, "LAD", 1) > 0)
        VendorMatch = "IN-HOUSE PRINTING"
    
    Case (InStr(1, det_desc, "star d", 1) > 0)
        VendorMatch = "STAR DIAGNOSIS - MERCEDES"
        
    Case (InStr(1, det_desc, "witech", 1) > 0)
        VendorMatch = "WITECH - CJD"
        
    Case (InStr(1, doc_desc, "RICOH", 1) > 0)
        VendorMatch = "RICOH USA INC"
    
    Case (InStr(1, det_desc, "edp", 1) > 0 Or (InStr(1, control1, "ADOBE", 1) > 0) Or (InStr(1, doc_desc, "EDP", 1) > 0))
        VendorMatch = "EDP CHARGES"
    
    Case (InStr(1, det_desc, "CVR", 1) > 0)
        VendorMatch = "COMPUTERIZED VEHICLE REGISTRATION"
    
    Case (InStr(1, det_desc, "CUDL", 1) > 0)
        VendorMatch = "CUDL CREDIT UNION DIRECT CORP"
     
    Case (InStr(1, doc_desc, "DMV", 1) > 0)
        VendorMatch = "DMV"
    
    Case (InStr(1, doc_desc, "VITU", 1) > 0)
        VendorMatch = "VITU"
    
    Case (InStr(1, det_desc, "SYS", 1) > 0 And InStr(1, det_desc, "FEE", 1) > 0)
        VendorMatch = "CHRYSLER SYSTEM FEE"
    
    Case (InStr(1, det_desc, "CDK DLR CAR", 1) > 0)
        VendorMatch = "CDK DLR CAR"
    
    Case (InStr(1, doc_desc, "cdk", 1) > 0) And (Not InStr(1, doc_desc, "dbs", 1) > 0)
        VendorMatch = "CDK GLOBAL LLC"
    
    Case (InStr(1, det_desc, "gm", 1) > 0) And (InStr(1, det_desc, "p", 1) > 0) And (InStr(1, det_desc, "w", 1) > 0)
        VendorMatch = "CDK GLOBAL LLC"
        
    Case (InStr(1, control2, "gm", 1) > 0) And (InStr(1, control2, "p", 1) > 0) And (InStr(1, control2, "w", 1) > 0)
        VendorMatch = "CDK GLOBAL LLC"
        
    Case (InStr(1, control1, "gm", 1) > 0) And (InStr(1, control1, "p", 1) > 0) And (InStr(1, control1, "w", 1) > 0)
        VendorMatch = "CDK GLOBAL LLC"

    
    'More Robust Searches based on which category the entry falls into
    
    Case (InStr(1, control1, "L", 1) = 1)
        VendorMatch = StoreMatch(control1, vendict, vendorlist)
    
    Case (InStr(1, doc_desc, "ftc", 1) > 0) 'FTC entries
        VendorMatch = FTCMatch(control1, control2, det_desc)
    
    Case (doc_desc = "STORE LAO PAYABLES ALLOCATION") 'LAO Payable entries
        VendorMatch = LAOMatch(det_desc, vendict, vendorlist, 0)
    
    Case ((InStr(1, doc_desc, "inter", 1) > 0) And (InStr(1, doc_desc, "co", 1) > 0)) 'Inter-company Billing entries
        VendorMatch = ICBMatch(control2, vendict, vendorlist) 'Check Control2 entry first
        
        If VendorMatch = "" Then 'Then check control1 if the first search fails
            VendorMatch = ICBMatch(control1, vendict, vendorlist)
        End If
        

    Case (doc_desc = "Check" And Not checklist Is Nothing) 'Check entries
        VendorMatch = CheckMatch(refnum, acct_date, checklist)
 
    
    VendorMatch = venname
    End Select
    
End Function

Function ICBMatch(stringin, vendict As Scripting.Dictionary, vendorlist As Worksheet)
    Dim srchval As String
    Dim elem As Variant
    Dim elem2 As Variant
    Dim venext() As String
    Dim namefields() As String
    Dim splitstr() As String
    Dim foundrow As Variant
    Dim venname As String
       
    
    srchval = Trim(stringin) 'Trim the search string
    venext = Split(" Inc, llc, ltd, co", ",") 'Build an array of the major company extensions for later checking
    namefields = Split("B:B,C:C,W:W", ",") 'Build an array with the column labels for all places that a company name might appear
    
    'If search string is an ICB code[of the form xxxx-xxxxxx-xxx], harvest the second piece as the search value
    If InStr(1, srchval, "-") > 0 Then
        splitstr = Split(CStr(srchval), "-")
        srchval = splitstr(1)
    End If
    
    'Determine what form the search string takes and act appropriately
    Select Case True
    
    'case 1 - search string blank: return the blank and pass up
    Case (srchval = "")
        ICBMatch = ""
        Exit Function
    
    'case 2 - search string is "acq": return Acquisition Expense as the vendor
    Case InStr(1, srchval, "acq", 1) > 0
        ICBMatch = "ACQUISITION EXPENSE"
        Exit Function
    
    'case 3 - search string is a number: check the vendor control number
    Case IsNumeric(srchval)
        foundrow = Application.Match(Val(srchval), vendorlist.Range("A:A"), 0)
    
    'case 4 - search string is a name: check the list of possible name fields
    Case Else
        If (srchval = "ACCUV") Then ICBMatch = "ACCUVANT INC": Exit Function 'catch the irritating accuvant code
        If (srchval = "LAD") Then ICBMatch = "LAD PRINT SHOP": Exit Function
        For Each elem In namefields 'Walk through each range of name fields, starting with Name 1, then Name 2, then Name 3
            foundrow = Application.Match(srchval, vendorlist.Range(elem), 0) 'Find it
            If Not IsError(foundrow) Then Exit For 'If foundrow is a number (meaning match found) then stop and move on
            For Each elem2 In venext 'Since a match wasn't found, add all the common extensions and see if something is found
                foundrow = Application.Match(srchval & elem2, vendorlist.Range(elem), 0) 'Find it with extension
                If Not IsError(foundrow) Then Exit For 'If match now found, exit both loops
            Next elem2
            
            If Not IsError(foundrow) Then Exit For 'If foundrow is a number (meaning match found) then stop and move on
        Next elem
    End Select
    
    'Lookup the value using the row found and return it
    If IsError(foundrow) Then 'If errored-out, bail
        venname = ""
    ElseIf (foundrow = 0) Then 'If other errored-out, bail
        venname = ""
    ElseIf IsNumeric(foundrow) Then 'If the Match found something, look it up in the vendor dictionary and return the vendor name
        venname = vendict(foundrow)
    Else 'If somehow everything else fails, bail
        venname = ""
    End If

    
    ICBMatch = venname

  
    
End Function

Function LAOMatch(stringin, vendict As Scripting.Dictionary, vendorlist As Worksheet, x As Integer)
    Dim elem As Variant
    Dim elem2 As Variant
    Dim venname As String
    Dim checkrange As Range
    Dim srchval As String
    Dim namefields() As String
    Dim venext() As String
    Dim foundrow As Variant
    Dim srchnoext() As String
    Dim andswitch() As String
    Dim ampswitch() As String
    
    srchval = Trim(stringin) 'Trim the string in prep for search
    namefields = Split("B:B,C:C,W:W", ",") 'Build an array of column ranges where names can appear
    venext = Split(" Inc, Llc, ltd, co", ",") 'Build an array of typical endings to company names
    
    If srchval = "" Then LAOMatch = "": Exit Function
    For Each elem In namefields 'For each range where names appear
    
        'This is the problem, it needs to be rewritten. Follow this link: http://www.ozgrid.com/forum/showthread.php?t=167487
       
        
        
        foundrow = Application.Match(srchval, vendorlist.Range(elem), 0) 'Look for the value in vendor list
        If Not IsError(foundrow) Then Exit For 'If match found, stop and move on
                 
        'if nothing is found, strip any extensions found and try another search
        For Each elem2 In venext 'For each extension
            If InStr(1, srchval, elem2, 1) > 0 Then 'If the search string has that ending
               srchnoext = Split(srchval, elem2, -1, 1) 'break extension off
               foundrow = Application.Match(srchnoext(0) & "*", vendorlist.Range(elem), 0) 'And search actual name with a wildcard
            End If
            If Not IsError(foundrow) Then Exit For 'If a match is found, move on
            Next elem2
        If Not IsError(foundrow) Then Exit For 'If a match is found, move on
        Set checkrange = Nothing
    Next elem
    
     'Lookup the value using the row found and return it
    If IsError(foundrow) Then 'If error, bail
        venname = ""
    ElseIf (foundrow = 0) Then 'If other error, bail
        venname = ""
    ElseIf IsNumeric(foundrow) Then 'If match found a row, look it up in the dictionary
        venname = vendict(foundrow)
    Else 'If somehow it still got goofed, bail
        venname = ""
    End If
    
    If venname = "" And x = 0 Then  'If the match failed (in any way), look for an and or ampersand and start over
        If InStr(1, srchval, "AND", 1) > 0 Then 'Check for 'And'
            andswitch = Split(srchval, "AND", 2, 1) 'Break it out
            venname = LAOMatch(andswitch(0) & "&" & andswitch(1), vendict, vendorlist, x + 1) 'Switch for ampersand and start over
        End If
    
        If InStr(1, srchval, "&", 1) > 0 Then 'Check for '&'
            ampswitch = Split(srchval, "&", 2, 1) 'Break it out
            venname = LAOMatch(ampswitch(0) & "&" & ampswitch(1), vendict, vendorlist, x + 1) 'Switch for 'and'
        End If
    End If
    

    LAOMatch = venname
End Function

Function JVMatch(stringin)
    srchval = Trim(stringin) 'Trim the search string
    
Select Case True 'Journal Vouchers fall into very few categories, return the correct vendor for each

    Case (InStr(1, srchval, "CUD", vbTextCompare) > 0)
        venname = "CUDL CREDIT UNION DIRECT CORP"

    Case (InStr(1, srchval, "ADP", vbTextCompare) > 0)
        venname = "ADP INC"
        
    Case (InStr(1, srchval, "CDK", vbTextCompare) > 0)
        venname = "CDK GLOBAL LLC"
        
    Case Else 'If it doesn't match anything else, bail
        venname = ""
End Select

JVMatch = venname

End Function
Function reuseoldven(refnum, acctdate, storenum, amt, oldsheet) As String
    For i = 2 To UBound(oldsheet, 1)
        If oldsheet(i, 8) = refnum And _
            oldsheet(i, 4) = acctdate And _
            oldsheet(i, 1) = storenum And _
            oldsheet(i, 12) = amt Then
                reuseoldven = oldsheet(i, 14)
                Exit Function
        End If
    Next i
    reuseoldven = ""
    
End Function


Function ReuseOldVen2(refnum, acctdate, storenum, amt, oldsheet As Worksheet) As String
    
    'Use advanced filters to list all checks with the right check number and store name (which should be unique) and snag it
    oldsheet.Range("A999998").Value = "Reference" 'Build tiny filter table way at the bottom of the check sheet
    oldsheet.Range("B999998").Value = "AcctgDate"
    oldsheet.Range("C999998").Value = "CoNo"
    oldsheet.Range("D999998").Value = "Amt"
    oldsheet.Range("A999999").Value = refnum 'Use reference number and check date for criteria
    oldsheet.Range("B999999").Value = CDate(acctdate)
    oldsheet.Range("C999999").Value = storenum
    oldsheet.Range("D999999").Value = amt
    
    'Execute an advanced filter using criteria that copies the (in theory) unique row from the check sheet
    oldsheet.Range("A1:AD999990").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        oldsheet.Range("A999998:D999999"), CopyToRange:=oldsheet.Range("A1000000"), Unique:= _
        False
    venname = oldsheet.Range("N1000001").Value 'grab the vendor name from the unique row
    oldsheet.Rows("999998:1000100").Delete 'delete the rows created for the advanced filter
    
    reuseoldven = venname
    
End Function

Function CheckMatch(refnum, checkdate, checklist As Worksheet)
    Dim venname As String
    
    'Use advanced filters to list all checks with the right check number and store name (which should be unique) and snag it
    checklist.Range("A499998").Value = "Reference" 'Build tiny filter table way at the bottom of the check sheet
    checklist.Range("B499998").Value = "AcctgDate"
    checklist.Range("A499999").Value = refnum 'Use reference number and check date for criteria
    checklist.Range("B499999").Value = CDate(checkdate)
    
    'Execute an advanced filter using criteria that copies the (in theory) unique row from the check sheet
    checklist.Range("A1:AD499990").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        checklist.Range("A499998:B499999"), CopyToRange:=checklist.Range("A500000"), Unique:= _
        False
    venname = checklist.Range("P500001").Value 'grab the vendor name from the unique row
    checklist.Rows("499998:500100").Delete 'delete the rows created for the advanced filter
    CheckMatch = venname 'pass up the vendor name

End Function

Function FTCMatch(control1, control2, det_desc)
    
    Select Case True 'FTC's fall into few categories that are easily predictable so grab the right one and pass up
    
    'Chrysler FTC's
    Case ((InStr(1, control1, "chry", 1) > 0) Or (InStr(1, control2, "chry", 1) > 0) Or (InStr(1, det_desc, "chry", 1) > 0))
        FTCMatch = "FTC - Chrysler": Exit Function
    
    'Ford FTC's and Lincoln FTC's
    Case ((InStr(1, control1, "ford", 1) > 0) Or (InStr(1, control2, "ford", 1) > 0) Or (InStr(1, det_desc, "ford", 1) > 0))
        If InStr(1, control1, "linc", 1) > 0 Then FTCMatch = "FTC - Lincoln": Exit Function
        FTCMatch = "FTC - Ford": Exit Function
    
    'Hyundai FTC's
    Case ((InStr(1, control1, "hyun", 1) > 0) Or (InStr(1, control2, "hyun", 1) > 0) Or (InStr(1, det_desc, "hyun", 1) > 0))
        FTCMatch = "FTC - Hyundai": Exit Function
    
    'Nissan FTC's
    Case ((InStr(1, control1, "niss", 1) > 0) Or (InStr(1, control2, "niss", 1) > 0) Or (InStr(1, det_desc, "niss", 1) > 0))
        FTCMatch = "FTC - Nissan": Exit Function
    
    'All other FTC's that can't be assigned
    Case Else: FTCMatch = "FTC - Undefined"
    
    End Select
End Function

Function StoreMatch(control1, vendict As Scripting.Dictionary, vendorlist As Worksheet)
    Dim srchval As String
    Dim foundrow As Variant
    Dim venname As String
    
    srchval = Trim(control1)
    foundrow = Application.Match(srchval, vendorlist.Range("A:A"), 0)
    
    If IsError(foundrow) Then 'If error, bail
        venname = ""
    ElseIf (foundrow = 0) Then 'If other error, bail
        venname = ""
    ElseIf IsNumeric(foundrow) Then 'If match found a row, look it up in the dictionary
        venname = vendict(foundrow)
    Else 'If somehow it still got goofed, bail
        venname = ""
    End If
    
    StoreMatch = venname
    
End Function

