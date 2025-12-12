Attribute VB_Name = "VerifyPMTData"
Sub SyncTechnicalFileToData()
    '==================================================================
    ' MODULE 1: Sync Technical File to Technical Data
    ' Verifies all items in Technical File exist in Technical Data
    ' Offers to copy missing items with their details
    '==================================================================
    
    Dim wb As Workbook
    Dim tfSheet As Worksheet    ' Technical File
    Dim tdSheet As Worksheet    ' Technical Data
    
    Dim tfItemIDCol As Long, tfAbbrCol As Long, tfNameCol As Long, tfRespCol As Long
    Dim tdItemIDCol As Long, tdAbbrCol As Long, tdNameCol As Long, tdRespCol As Long
    
    Dim tfLastRow As Long, tdLastRow As Long
    Dim i As Long, j As Long
    Dim itemID As String
    Dim foundInTD As Boolean
    Dim missingItems As Collection
    Dim insertRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Set workbook and sheets
    Set wb = ThisWorkbook
    Set tfSheet = wb.Worksheets("Technical File")
    Set tdSheet = wb.Worksheets("Technical Data")
    
    ' Find columns in Technical File (row 3)
    tfItemIDCol = FindColumn(tfSheet, 3, "ITEM ID")
    tfAbbrCol = FindColumn(tfSheet, 3, "ABBREVIATION")
    tfNameCol = FindColumn(tfSheet, 3, "NAME")
    tfRespCol = FindColumn(tfSheet, 3, "RESPONSIBLE")
    
    If tfItemIDCol = 0 Then
        MsgBox "Column 'Item ID' not found in Technical File (row 3).", vbExclamation
        Exit Sub
    End If
    
    ' Find columns in Technical Data (row 3)
    tdItemIDCol = FindColumn(tdSheet, 3, "ITEM ID")
    tdAbbrCol = FindColumn(tdSheet, 3, "ABBREVIATION")
    tdNameCol = FindColumn(tdSheet, 3, "NAME")
    tdRespCol = FindColumn(tdSheet, 3, "RESPONSIBLE")
    
    If tdItemIDCol = 0 Then
        MsgBox "Column 'Item ID' not found in Technical Data (row 3).", vbExclamation
        Exit Sub
    End If
    
    ' Get last rows - search downward from row 7 to find last non-empty cell
    tfLastRow = 7
    For i = 7 To 1000  ' Check up to row 1000
        If Trim(tfSheet.Cells(i, tfItemIDCol).Value) <> "" Then
            tfLastRow = i
        End If
    Next i
    
    tdLastRow = 7
    For i = 7 To 1000  ' Check up to row 1000
        If Trim(tdSheet.Cells(i, tdItemIDCol).Value) <> "" Then
            tdLastRow = i
        End If
    Next i
    
    ' Find missing items
    Set missingItems = New Collection
    
    For i = 7 To tfLastRow  ' Start from row 7
        itemID = Trim(tfSheet.Cells(i, tfItemIDCol).Value)
        
        If itemID <> "" Then
            ' Check if item exists in Technical Data
            foundInTD = False
            For j = 7 To tdLastRow
                If Trim(tdSheet.Cells(j, tdItemIDCol).Value) = itemID Then
                    foundInTD = True
                    Exit For
                End If
            Next j
            
            If Not foundInTD Then
                On Error Resume Next
                missingItems.Add Array(i, itemID), CStr(itemID)  ' Store row and itemID
                On Error GoTo ErrorHandler
            End If
        End If
    Next i
    
    ' Report findings
    If missingItems.count = 0 Then
        MsgBox "All items from Technical File exist in Technical Data!" & vbCrLf & vbCrLf & _
               "Total items checked: " & (tfLastRow - 6), vbInformation, "Sync Complete"
        Exit Sub
    End If
    
    ' Ask if user wants to copy missing items
    Dim response As VbMsgBoxResult
    Dim missingList As String
    missingList = missingItems.count & " item(s) found in Technical File but missing in Technical Data:" & vbCrLf & vbCrLf
    
    Dim item As Variant
    Dim count As Long
    count = 0
    For Each item In missingItems
        count = count + 1
        missingList = missingList & "- " & item(1) & vbCrLf
        If count >= 10 And missingItems.count > 10 Then
            missingList = missingList & "... and " & (missingItems.count - 10) & " more" & vbCrLf
            Exit For
        End If
    Next item
    
    missingList = missingList & vbCrLf & "Copy these items to Technical Data?"
    
    response = MsgBox(missingList, vbYesNo + vbQuestion, "Missing Items Found")
    
    If response = vbNo Then
        MsgBox "Sync cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Copy missing items in correct order
    Dim itemsAdded As Long
    itemsAdded = 0
    
    For Each item In missingItems
        Dim tfRow As Long
        tfRow = item(0)  ' Original row in Technical File
        itemID = item(1)
        
        ' Find correct insert position in Technical Data (maintain order)
        insertRow = FindInsertPosition(tfSheet, tdSheet, tfRow, tfItemIDCol, tdItemIDCol)
        
        ' Insert new row
        tdSheet.Rows(insertRow).Insert Shift:=xlDown
        tdLastRow = tdLastRow + 1
        
        ' Copy data
        If tdItemIDCol > 0 Then tdSheet.Cells(insertRow, tdItemIDCol).Value = tfSheet.Cells(tfRow, tfItemIDCol).Value
        If tdAbbrCol > 0 And tfAbbrCol > 0 Then tdSheet.Cells(insertRow, tdAbbrCol).Value = tfSheet.Cells(tfRow, tfAbbrCol).Value
        If tdNameCol > 0 And tfNameCol > 0 Then tdSheet.Cells(insertRow, tdNameCol).Value = tfSheet.Cells(tfRow, tfNameCol).Value
        If tdRespCol > 0 And tfRespCol > 0 Then tdSheet.Cells(insertRow, tdRespCol).Value = tfSheet.Cells(tfRow, tfRespCol).Value
        
        ' Highlight new row
        tdSheet.Rows(insertRow).Interior.Color = RGB(144, 238, 144)  ' Light green
        
        itemsAdded = itemsAdded + 1
    Next item
    
    ' Success message
    MsgBox itemsAdded & " item(s) successfully copied to Technical Data (highlighted in GREEN)." & vbCrLf & vbCrLf & _
           "Items were inserted in the correct order.", vbInformation, "Sync Complete"
    
    wb.Save
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub


Sub VerifyTechnicalSheets()
    '==================================================================
    ' MODULE 2: Comprehensive Verification with Highlighting
    ' Checks that both sheets are properly synchronized
    ' Verifies order, presence, and matching details
    ' Highlights mismatches and missing items in ORANGE
    '==================================================================
    
    Dim wb As Workbook
    Dim tfSheet As Worksheet    ' Technical File
    Dim tdSheet As Worksheet    ' Technical Data
    
    Dim tfItemIDCol As Long, tfAbbrCol As Long, tfNameCol As Long, tfRespCol As Long
    Dim tdItemIDCol As Long, tdAbbrCol As Long, tdNameCol As Long, tdRespCol As Long
    Dim tdTechFileCol As Long  ' Technical File (Y/N) column
    
    Dim tfLastRow As Long, tdLastRow As Long
    Dim i As Long, j As Long
    Dim itemID As String
    
    Dim errors As Collection
    Dim errorMsg As String
    Dim errorCount As Long
    Dim issuesFound As Long
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set tfSheet = wb.Worksheets("Technical File")
    Set tdSheet = wb.Worksheets("Technical Data")
    Set errors = New Collection
    
    ' Find columns in Technical File (row 3)
    tfItemIDCol = FindColumn(tfSheet, 3, "ITEM ID")
    tfAbbrCol = FindColumn(tfSheet, 3, "ABBREVIATION")
    tfNameCol = FindColumn(tfSheet, 3, "NAME")
    tfRespCol = FindColumn(tfSheet, 3, "RESPONSIBLE")
    
    ' Find columns in Technical Data (row 3)
    tdItemIDCol = FindColumn(tdSheet, 3, "ITEM ID")
    tdAbbrCol = FindColumn(tdSheet, 3, "ABBREVIATION")
    tdNameCol = FindColumn(tdSheet, 3, "NAME")
    tdRespCol = FindColumn(tdSheet, 3, "RESPONSIBLE")
    tdTechFileCol = FindColumn(tdSheet, 3, "TECHNICAL FILE (Y/N)")
    
    If tfItemIDCol = 0 Or tdItemIDCol = 0 Then
        MsgBox "Item ID column not found in one or both sheets.", vbExclamation
        Exit Sub
    End If
    
    ' Get last rows - search downward from row 7
    tfLastRow = 7
    For i = 7 To 1000
        If Trim(tfSheet.Cells(i, tfItemIDCol).Value) <> "" Then
            tfLastRow = i
        End If
    Next i
    
    tdLastRow = 7
    For i = 7 To 1000
        If Trim(tdSheet.Cells(i, tdItemIDCol).Value) <> "" Then
            tdLastRow = i
        End If
    Next i
    
    ' Clear previous highlighting in Technical Data (rows 7 onwards)
    If tdLastRow >= 7 Then
        tdSheet.Rows("7:" & tdLastRow).Interior.ColorIndex = xlNone
    End If
    
    issuesFound = 0
    
    ' CHECK 1: All items in Technical File must be in Technical Data
    For i = 7 To tfLastRow
        itemID = Trim(tfSheet.Cells(i, tfItemIDCol).Value)
        If itemID <> "" Then
            Dim foundInTD As Boolean
            Dim tdRowFound As Long
            foundInTD = False
            tdRowFound = 0
            
            For j = 7 To tdLastRow
                If Trim(tdSheet.Cells(j, tdItemIDCol).Value) = itemID Then
                    foundInTD = True
                    tdRowFound = j
                    
                    ' Verify matching details and highlight mismatches
                    Dim tfAbbr As String, tfName As String, tfResp As String
                    Dim tdAbbr As String, tdName As String, tdResp As String
                    Dim hasMismatch As Boolean
                    hasMismatch = False
                    
                    tfAbbr = Trim(tfSheet.Cells(i, tfAbbrCol).Value)
                    tdAbbr = Trim(tdSheet.Cells(j, tdAbbrCol).Value)
                    
                    tfName = Trim(tfSheet.Cells(i, tfNameCol).Value)
                    tdName = Trim(tdSheet.Cells(j, tdNameCol).Value)
                    
                    tfResp = Trim(tfSheet.Cells(i, tfRespCol).Value)
                    tdResp = Trim(tdSheet.Cells(j, tdRespCol).Value)
                    
                    ' Check Abbreviation mismatch
                    If tfAbbrCol > 0 And tdAbbrCol > 0 And tfAbbr <> tdAbbr Then
                        errors.Add "Item " & itemID & " (TD row " & j & "): Abbreviation mismatch (TF: '" & tfAbbr & "' vs TD: '" & tdAbbr & "')"
                        tdSheet.Cells(j, tdAbbrCol).Interior.Color = RGB(255, 200, 0)  ' Orange
                        hasMismatch = True
                        issuesFound = issuesFound + 1
                    End If
                    
                    ' Check Name mismatch
                    If tfNameCol > 0 And tdNameCol > 0 And tfName <> tdName Then
                        errors.Add "Item " & itemID & " (TD row " & j & "): Name mismatch (TF: '" & tfName & "' vs TD: '" & tdName & "')"
                        tdSheet.Cells(j, tdNameCol).Interior.Color = RGB(255, 200, 0)  ' Orange
                        hasMismatch = True
                        issuesFound = issuesFound + 1
                    End If
                    
                    ' Check Responsible mismatch
                    If tfRespCol > 0 And tdRespCol > 0 And tfResp <> tdResp Then
                        errors.Add "Item " & itemID & " (TD row " & j & "): Responsible mismatch (TF: '" & tfResp & "' vs TD: '" & tdResp & "')"
                        tdSheet.Cells(j, tdRespCol).Interior.Color = RGB(255, 200, 0)  ' Orange
                        hasMismatch = True
                        issuesFound = issuesFound + 1
                    End If
                    
                    ' Highlight Item ID if there are any mismatches in this row
                    If hasMismatch Then
                        tdSheet.Cells(j, tdItemIDCol).Interior.Color = RGB(255, 200, 0)  ' Orange
                    End If
                    
                    Exit For
                End If
            Next j
            
            If Not foundInTD Then
                errors.Add "Item " & itemID & " (TF row " & i & "): In Technical File but MISSING in Technical Data"
                issuesFound = issuesFound + 1
                ' Can't highlight in TD as it doesn't exist there
            End If
        End If
    Next i
    
    ' CHECK 2: All items in Technical Data with "Y" must be in Technical File
    If tdTechFileCol > 0 Then
        For i = 7 To tdLastRow
            itemID = Trim(tdSheet.Cells(i, tdItemIDCol).Value)
            Dim techFileFlag As String
            techFileFlag = UCase(Trim(tdSheet.Cells(i, tdTechFileCol).Value))
            
            If itemID <> "" And (techFileFlag = "Y" Or techFileFlag = "YES") Then
                Dim foundInTF As Boolean
                foundInTF = False
                
                For j = 7 To tfLastRow
                    If Trim(tfSheet.Cells(j, tfItemIDCol).Value) = itemID Then
                        foundInTF = True
                        Exit For
                    End If
                Next j
                
                If Not foundInTF Then
                    errors.Add "Item " & itemID & " (TD row " & i & "): Marked 'Y' in Technical Data but MISSING in Technical File"
                    ' Highlight entire row in orange
                    tdSheet.Rows(i).Interior.Color = RGB(255, 200, 0)  ' Orange
                    issuesFound = issuesFound + 1
                End If
            End If
        Next i
    End If
    
    ' CHECK 3: Verify order is the same
    Dim tfIndex As Long, tdIndex As Long
    tfIndex = 7
    tdIndex = 7
    
    Dim orderErrors As Long
    orderErrors = 0
    
    Do While tfIndex <= tfLastRow And tdIndex <= tdLastRow And orderErrors < 5
        Dim tfItem As String, tdItem As String
        tfItem = Trim(tfSheet.Cells(tfIndex, tfItemIDCol).Value)
        tdItem = Trim(tdSheet.Cells(tdIndex, tdItemIDCol).Value)
        
        ' Skip empty rows
        If tfItem = "" Then
            tfIndex = tfIndex + 1
        ElseIf tdItem = "" Then
            tdIndex = tdIndex + 1
        ElseIf tfItem <> tdItem Then
            errors.Add "Order mismatch at TF row " & tfIndex & " / TD row " & tdIndex & _
                      ": TF has '" & tfItem & "', TD has '" & tdItem & "'"
            ' Highlight the mismatched row in Technical Data
            If tdSheet.Rows(tdIndex).Interior.Color <> RGB(255, 200, 0) Then
                tdSheet.Cells(tdIndex, tdItemIDCol).Interior.Color = RGB(255, 220, 150)  ' Light orange for order issues
            End If
            orderErrors = orderErrors + 1
            issuesFound = issuesFound + 1
            tfIndex = tfIndex + 1
            tdIndex = tdIndex + 1
        Else
            tfIndex = tfIndex + 1
            tdIndex = tdIndex + 1
        End If
    Loop
    
    ' Report results
    If errors.count = 0 Then
        MsgBox "✓ VERIFICATION PASSED!" & vbCrLf & vbCrLf & _
               "Both sheets are properly synchronized:" & vbCrLf & _
               "- All items present in both sheets" & vbCrLf & _
               "- Order matches" & vbCrLf & _
               "- All details match" & vbCrLf & vbCrLf & _
               "Items checked: " & (tfLastRow - 6), vbInformation, "Verification Complete"
    Else
        errorMsg = "⚠ VERIFICATION FAILED - " & errors.count & " issue(s) found:" & vbCrLf & vbCrLf
        
        errorCount = 0
        For i = 1 To errors.count
            errorCount = errorCount + 1
            errorMsg = errorMsg & errorCount & ". " & errors(i) & vbCrLf
            
            If errorCount >= 15 And errors.count > 15 Then
                errorMsg = errorMsg & vbCrLf & "... and " & (errors.count - 15) & " more issues"
                Exit For
            End If
        Next i
        
        errorMsg = errorMsg & vbCrLf & vbCrLf & _
                   "Mismatches and missing items are highlighted in ORANGE in Technical Data sheet."
        
        MsgBox errorMsg, vbExclamation, "Verification Issues Found"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub


Sub ClearTechnicalDataHighlighting()
    '==================================================================
    ' HELPER: Clear all highlighting from Technical Data
    '==================================================================
    Dim wb As Workbook
    Dim tdSheet As Worksheet
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set tdSheet = wb.Worksheets("Technical Data")
    
    ' Find last row
    lastRow = 7
    Dim tdItemIDCol As Long
    tdItemIDCol = FindColumn(tdSheet, 3, "ITEM ID")
    
    If tdItemIDCol > 0 Then
        For i = 7 To 1000
            If Trim(tdSheet.Cells(i, tdItemIDCol).Value) <> "" Then
                lastRow = i
            End If
        Next i
    End If
    
    ' Clear highlighting
    If lastRow >= 7 Then
        tdSheet.Rows("7:" & lastRow).Interior.ColorIndex = xlNone
        MsgBox "Highlighting cleared from Technical Data sheet.", vbInformation, "Clear Complete"
    Else
        MsgBox "No data found to clear.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub


'==================================================================
' HELPER FUNCTIONS
'==================================================================

Function FindColumn(ws As Worksheet, headerRow As Long, columnName As String) As Long
    ' Find column by header name (case-insensitive)
    Dim i As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If UCase(Trim(ws.Cells(headerRow, i).Value)) = UCase(Trim(columnName)) Then
            FindColumn = i
            Exit Function
        End If
    Next i
    
    FindColumn = 0  ' Not found
End Function


Function FindInsertPosition(tfSheet As Worksheet, tdSheet As Worksheet, _
                           tfRow As Long, tfItemIDCol As Long, tdItemIDCol As Long) As Long
    ' Find correct position to insert item in Technical Data to maintain order
    Dim i As Long
    Dim tfIndex As Long
    Dim tdLastRow As Long
    
    tdLastRow = 7
    For i = 7 To 1000
        If Trim(tdSheet.Cells(i, tdItemIDCol).Value) <> "" Then
            tdLastRow = i
        End If
    Next i
    
    ' Find the next item in Technical File after tfRow
    Dim nextItemID As String
    nextItemID = ""
    
    Dim tfLastRow As Long
    tfLastRow = 7
    For i = 7 To 1000
        If Trim(tfSheet.Cells(i, tfItemIDCol).Value) <> "" Then
            tfLastRow = i
        End If
    Next i
    
    For i = tfRow + 1 To tfLastRow
        nextItemID = Trim(tfSheet.Cells(i, tfItemIDCol).Value)
        If nextItemID <> "" Then Exit For
    Next i
    
    ' If no next item, insert at end
    If nextItemID = "" Then
        FindInsertPosition = tdLastRow + 1
        Exit Function
    End If
    
    ' Find where that next item is in Technical Data
    For i = 7 To tdLastRow
        If Trim(tdSheet.Cells(i, tdItemIDCol).Value) = nextItemID Then
            FindInsertPosition = i
            Exit Function
        End If
    Next i
    
    ' If next item not found in TD, insert at end
    FindInsertPosition = tdLastRow + 1
End Function
