Attribute VB_Name = "Module2"
Sub VerifyOptionalityHeaders()
    '==================================================================
    ' Verify Optionality Headers Match
    ' Checks that headers in Optionality Points List (D1:X4) match
    ' headers in Technical File (named range OPTIONALITY_HEADERS)
    ' in the same order
    '==================================================================
    
    Dim wb As Workbook
    Dim oplSheet As Worksheet
    Dim tfSheet As Worksheet
    Dim oplRange As Range
    Dim optHeadersRange As Range
    
    Dim oplHeaders As String
    Dim tfHeaders As String
    Dim i As Long
    Dim errors As Collection
    Dim mismatchFound As Boolean
    Dim totalCells As Long
    Dim matchCount As Long
    Dim mismatchCount As Long
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set oplSheet = wb.Worksheets("Optionality Points List")
    Set tfSheet = wb.Worksheets("Technical File")
    Set errors = New Collection
    
    ' Get the ranges
    Set oplRange = oplSheet.Range("D1:X4")
    
    ' Check if named range exists
    On Error Resume Next
    Set optHeadersRange = wb.Names("OPTIONALITY_HEADERS").RefersToRange
    On Error GoTo ErrorHandler
    
    If optHeadersRange Is Nothing Then
        MsgBox "Named range 'OPTIONALITY_HEADERS' not found in workbook." & vbCrLf & vbCrLf & _
               "Please create this named range in the Technical File sheet.", _
               vbExclamation, "Named Range Not Found"
        Exit Sub
    End If
    
    ' Verify both ranges have same dimensions
    If oplRange.Rows.count <> optHeadersRange.Rows.count Then
        errors.Add "Row count mismatch: OPL has " & oplRange.Rows.count & " rows, OPTIONALITY_HEADERS has " & optHeadersRange.Rows.count & " rows"
    End If
    
    If oplRange.Columns.count <> optHeadersRange.Columns.count Then
        errors.Add "Column count mismatch: OPL has " & oplRange.Columns.count & " columns, OPTIONALITY_HEADERS has " & optHeadersRange.Columns.count & " columns"
    End If
    
    ' If dimensions don't match, report and exit
    If errors.count > 0 Then
        MsgBox "? DIMENSION MISMATCH" & vbCrLf & vbCrLf & _
               "Optionality Points List (D1:X4): " & oplRange.Rows.count & " rows × " & oplRange.Columns.count & " columns" & vbCrLf & _
               "OPTIONALITY_HEADERS: " & optHeadersRange.Rows.count & " rows × " & optHeadersRange.Columns.count & " columns" & vbCrLf & vbCrLf & _
               "Ranges must have the same dimensions.", _
               vbExclamation, "Verification Failed"
        Exit Sub
    End If
    
    mismatchFound = False
    matchCount = 0
    mismatchCount = 0
    totalCells = oplRange.Rows.count * oplRange.Columns.count
    
    ' Compare cell by cell
    Dim row As Long, col As Long
    Dim oplCell As Range, tfCell As Range
    
    For row = 1 To oplRange.Rows.count
        For col = 1 To oplRange.Columns.count
            Set oplCell = oplRange.Cells(row, col)
            Set tfCell = optHeadersRange.Cells(row, col)
            
            Dim oplValue As String, tfValue As String
            oplValue = Trim(oplCell.Value)
            tfValue = Trim(tfCell.Value)
            
            If oplValue <> tfValue Then
                ' Record the mismatch
                errors.Add "Position [Row " & row & ", Col " & col & "]: " & _
                          "OPL has '" & oplValue & "', " & _
                          "TF has '" & tfValue & "'"
                
                mismatchFound = True
                mismatchCount = mismatchCount + 1
            Else
                matchCount = matchCount + 1
            End If
        Next col
    Next row
    
    ' Report results
    If Not mismatchFound Then
        MsgBox "? VERIFICATION PASSED!" & vbCrLf & vbCrLf & _
               "All headers match in the correct order:" & vbCrLf & vbCrLf & _
               "Summary:" & vbCrLf & _
               "  • Rows checked: " & oplRange.Rows.count & vbCrLf & _
               "  • Columns checked: " & oplRange.Columns.count & vbCrLf & _
               "  • Total cells verified: " & totalCells & vbCrLf & _
               "  • Matches: " & matchCount & vbCrLf & _
               "  • Differences: " & mismatchCount, _
               vbInformation, "Verification Complete"
    Else
        Dim errorMsg As String
        errorMsg = "? VERIFICATION FAILED" & vbCrLf & vbCrLf & _
                   "Summary:" & vbCrLf & _
                   "  • Total cells checked: " & totalCells & vbCrLf & _
                   "  • Matches: " & matchCount & vbCrLf & _
                   "  • Differences found: " & mismatchCount & vbCrLf & vbCrLf & _
                   "Details of mismatches:" & vbCrLf
        
        Dim errorCount As Long
        errorCount = 0
        For i = 1 To errors.count
            errorCount = errorCount + 1
            errorMsg = errorMsg & errorCount & ". " & errors(i) & vbCrLf
            
            If errorCount >= 10 And errors.count > 10 Then
                errorMsg = errorMsg & vbCrLf & "... and " & (errors.count - 10) & " more mismatches"
                Exit For
            End If
        Next i
        
        MsgBox errorMsg, vbExclamation, "Verification Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub


Sub ShowDetailedComparison()
    '==================================================================
    ' Show detailed side-by-side comparison in Immediate Window
    '==================================================================
    
    Dim wb As Workbook
    Dim oplSheet As Worksheet
    Dim oplRange As Range
    Dim optHeadersRange As Range
    Dim row As Long, col As Long
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set oplSheet = wb.Worksheets("Optionality Points List")
    Set oplRange = oplSheet.Range("D1:X4")
    
    On Error Resume Next
    Set optHeadersRange = wb.Names("OPTIONALITY_HEADERS").RefersToRange
    On Error GoTo ErrorHandler
    
    If optHeadersRange Is Nothing Then
        MsgBox "Named range 'OPTIONALITY_HEADERS' not found.", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "=== OPTIONALITY HEADERS COMPARISON ==="
    Debug.Print ""
    
    For row = 1 To oplRange.Rows.count
        Debug.Print "Row " & row & ":"
        For col = 1 To oplRange.Columns.count
            Dim oplVal As String, tfVal As String
            oplVal = Trim(oplRange.Cells(row, col).Value)
            tfVal = Trim(optHeadersRange.Cells(row, col).Value)
            
            If oplVal = tfVal Then
                Debug.Print "  Col " & col & ": ? '" & oplVal & "'"
            Else
                Debug.Print "  Col " & col & ": ? OPL='" & oplVal & "' vs TF='" & tfVal & "'"
            End If
        Next col
        Debug.Print ""
    Next row
    
    MsgBox "Detailed comparison printed to Immediate Window (Ctrl+G)", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Sub VerifyOptionalityPointColumns()
    '==================================================================
    ' Verify Optionality Point Columns
    ' Checks that all optionality points from OPL have columns in TF
    ' starting at OPTIONALITY_START position
    ' Only processes rows that have a title in column C
    '==================================================================
    
    Dim wb As Workbook
    Dim oplSheet As Worksheet
    Dim tfSheet As Worksheet
    
    Dim optStartRange As Range
    Dim startCol As Long
    Dim oplLastRow As Long
    Dim tfRow4LastCol As Long
    
    Dim oplID As String
    Dim oplTitle As String
    Dim foundInTF As Boolean
    Dim missingIDs As Collection
    Dim i As Long, j As Long
    
    Dim totalOPL As Long
    Dim foundCount As Long
    Dim missingCount As Long
    Dim skippedCount As Long
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set oplSheet = wb.Worksheets("Optionality Points List")
    Set tfSheet = wb.Worksheets("Technical File")
    Set missingIDs = New Collection
    
    ' Check if OPTIONALITY_START named range exists
    On Error Resume Next
    Set optStartRange = wb.Names("OPTIONALITY_START").RefersToRange
    On Error GoTo ErrorHandler
    
    If optStartRange Is Nothing Then
        MsgBox "Named range 'OPTIONALITY_START' not found in workbook." & vbCrLf & vbCrLf & _
               "Please create this named range in the Technical File sheet to mark where optionality columns start.", _
               vbExclamation, "Named Range Not Found"
        Exit Sub
    End If
    
    ' Get starting column for optionality in TF
    startCol = optStartRange.Column
    
    ' Find last row with data in OPL column A (first column)
    oplLastRow = 5
    For i = 5 To 1000
        If Trim(oplSheet.Cells(i, 1).Value) <> "" Then
            oplLastRow = i
        End If
    Next i
    
    ' Find last column in TF row 4 (to know where to search)
    tfRow4LastCol = startCol
    For i = startCol To tfSheet.Columns.count
        If Trim(tfSheet.Cells(4, i).Value) = "" Then
            tfRow4LastCol = i - 1
            Exit For
        End If
    Next i
    
    totalOPL = 0
    foundCount = 0
    missingCount = 0
    skippedCount = 0
    
    ' Check each optionality point ID in OPL
    For i = 5 To oplLastRow
        oplID = Trim(oplSheet.Cells(i, 1).Value)
        oplTitle = Trim(oplSheet.Cells(i, 3).Value)  ' Column C is the title
        
        ' Only process rows that have both an ID and a title
        If oplID <> "" Then
            If oplTitle <> "" Then
                ' Row has a title - process it
                totalOPL = totalOPL + 1
                foundInTF = False
                
                ' Search for this ID in TF row 4, starting from OPTIONALITY_START
                For j = startCol To tfRow4LastCol
                    If Trim(tfSheet.Cells(4, j).Value) = oplID Then
                        foundInTF = True
                        foundCount = foundCount + 1
                        Exit For
                    End If
                Next j
                
                ' If not found, add to missing collection
                If Not foundInTF Then
                    On Error Resume Next
                    missingIDs.Add Array(oplID, i, oplTitle), CStr(oplID)  ' Store ID, row number, and title
                    On Error GoTo ErrorHandler
                    missingCount = missingCount + 1
                End If
            Else
                ' Row has ID but no title - skip it
                skippedCount = skippedCount + 1
            End If
        End If
    Next i
    
    ' Report results
    If missingCount = 0 Then
        Dim successMsg As String
        successMsg = "? VERIFICATION PASSED!" & vbCrLf & vbCrLf & _
                     "All optionality points with titles from OPL have columns in Technical File." & vbCrLf & vbCrLf & _
                     "Summary:" & vbCrLf & _
                     "  • Total optionality points checked: " & totalOPL & vbCrLf & _
                     "  • Found in Technical File: " & foundCount & vbCrLf & _
                     "  • Missing: " & missingCount
        
        If skippedCount > 0 Then
            successMsg = successMsg & vbCrLf & vbCrLf & _
                        "Note: " & skippedCount & " row(s) skipped (no title in column C)"
        End If
        
        MsgBox successMsg, vbInformation, "Verification Complete"
        Exit Sub
    End If
    
    ' Missing columns found - ask user what to do
    Dim missingMsg As String
    missingMsg = "? VERIFICATION FAILED" & vbCrLf & vbCrLf & _
                 "Summary:" & vbCrLf & _
                 "  • Total optionality points with titles: " & totalOPL & vbCrLf & _
                 "  • Found in Technical File: " & foundCount & vbCrLf & _
                 "  • Missing in Technical File: " & missingCount
    
    If skippedCount > 0 Then
        missingMsg = missingMsg & vbCrLf & _
                     "  • Rows skipped (no title): " & skippedCount
    End If
    
    missingMsg = missingMsg & vbCrLf & vbCrLf & "Missing optionality point IDs:" & vbCrLf
    
    Dim count As Long
    count = 0
    Dim item As Variant
    For Each item In missingIDs
        count = count + 1
        missingMsg = missingMsg & "  " & count & ". " & item(0) & " - """ & item(2) & """ (OPL row " & item(1) & ")" & vbCrLf
        If count >= 15 And missingIDs.count > 15 Then
            missingMsg = missingMsg & "  ... and " & (missingIDs.count - 15) & " more" & vbCrLf
            Exit For
        End If
    Next item
    
    missingMsg = missingMsg & vbCrLf & "Do you want to insert missing columns in Technical File?"
    
    Dim response As VbMsgBoxResult
    response = MsgBox(missingMsg, vbYesNo + vbQuestion, "Missing Columns Found")
    
    If response = vbNo Then
        MsgBox "No columns were added to Technical File.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Insert missing columns
    Dim columnsAdded As Long
    columnsAdded = 0
    Dim insertCol As Long
    
    ' Insert at the end of existing optionality columns
    insertCol = tfRow4LastCol + 1
    
    For Each item In missingIDs
        oplID = item(0)
        
        ' Insert new column
        tfSheet.Columns(insertCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        ' Add the ID in row 4
        tfSheet.Cells(4, insertCol).Value = oplID
        
        columnsAdded = columnsAdded + 1
        insertCol = insertCol + 1
    Next item
    
    ' Success message
    MsgBox "? COLUMNS ADDED SUCCESSFULLY!" & vbCrLf & vbCrLf & _
           columnsAdded & " new column(s) inserted in Technical File." & vbCrLf & vbCrLf & _
           "New columns added starting at column " & Split(Cells(1, tfRow4LastCol + 1).Address, "$")(1) & vbCrLf & vbCrLf & _
           "Note: You may need to add headers in rows 1-3 and copy formulas/formatting.", _
           vbInformation, "Columns Added"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

