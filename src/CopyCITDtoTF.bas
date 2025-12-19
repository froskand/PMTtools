Attribute VB_Name = "CopyCITDtoTF"
' ========================================
' Copy CI to Technical File Functionality
' ========================================

' Function AddCopyToTechFileContextMenu()
' Purpose: Add "Copy CI to Technical file" option to Technical Data sheet context menu
' Key actions:
' - Add control to Row context menu
' Code flow Summary:
' - Obtains Context Menu, if fails informs
' - Add control, if fails informs
' - Configures Control

Sub AddCopyToTechFileContextMenu()
    Dim contextMenu As CommandBar
    Dim control As CommandBarControl
    
    ' Obtain Context Menu
    On Error Resume Next
    Set contextMenu = Application.CommandBars("Row")
    If contextMenu Is Nothing Then
        MsgBox "Cannot access the row context menu.", vbCritical
        Exit Sub
    End If
    
    ' Add Control
    Set control = contextMenu.Controls.Add(Type:=msoControlButton)
    If control Is Nothing Then
        MsgBox "Failed to add menu item.", vbCritical
        Exit Sub
    End If
    
    'Configure Control
    With control
        .Caption = "Copy CI to Technical file"
        .OnAction = "CopyRowToTechnicalFile"
        .BeginGroup = True
    End With
End Sub


' Function RemoveCopyToTechFileContextMenu()
' Purpose: Remove "Copy CI to Technical file" option from context menu
' Key actions:
' - Remove control
' Code flow Summary:
' - Obtains Context Menu, if fails exit
' - Find and remove menu item

Sub RemoveCopyToTechFileContextMenu()
    Dim ctrl As CommandBarControl
    Dim contextMenu As CommandBar

    Set contextMenu = Application.CommandBars("Row")
    If contextMenu Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ctrl In contextMenu.Controls
        If ctrl.Caption = "Copy CI to Technical file" Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub


' Function CopyRowToTechnicalFile()
' Purpose: Copy selected row from Technical Data sheet to Technical File sheet
'          Inserts row in correct position based on Item ID ordering
' Key actions:
' - Validate source sheet and row
' - Find columns in both sheets
' - Determine insert position based on Item ID
' - Insert row and column using InsertMatrixRowAndColumn
' - Copy data from source to target
' Code flow Summary:
' - Check if current sheet is Technical Data
' - Get selected row and validate
' - Find column indices in both sheets
' - Extract data from source row
' - Find insert position in Technical File based on Item ID
' - Call InsertMatrixRowAndColumn to insert at correct position
' - Copy data to new row

Sub CopyRowToTechnicalFile()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim sourceItemID As String
    Dim lastRow As Long
    Dim i As Long
    Dim insertPosition As Long
    
    ' Column mappings
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    
    ' Source data variables
    Dim identDate As Variant, location As Variant, itemID As Variant
    Dim abbreviation As Variant, name As Variant, responsible As Variant, version As Variant
    
    ' Check if we're on Technical Data sheet
    On Error Resume Next
    Set sourceSheet = ActiveSheet
    On Error GoTo 0
    
    If sourceSheet Is Nothing Then
        MsgBox "Cannot determine active sheet.", vbCritical
        Exit Sub
    End If
    
    ' Verify we're on Technical Data sheet (adjust name as needed)
    If sourceSheet.name <> "Technical Data" Then
        MsgBox "This function must be run from the Technical Data sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Get selected row
    sourceRow = Selection.row
    If sourceRow < 7 Then ' Assuming row 1-6 is header
        MsgBox "Please select a data row (not header).", vbExclamation
        Exit Sub
    End If
    
    ' Get target sheet
    On Error Resume Next
    Set targetSheet = Worksheets("Technical File")
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "Technical File sheet not found.", vbCritical
        Exit Sub
    End If
    
    ' Find column indices in source sheet (Technical Data)
    Dim srcIdentDateCol As Long, srcLocationCol As Long, srcItemIDCol As Long
    Dim srcAbbrevCol As Long, srcNameCol As Long, srcRespCol As Long, srcVersionCol As Long
    
    srcIdentDateCol = FindColumn(sourceSheet, "Identified Date")
    srcLocationCol = FindColumn(sourceSheet, "Location")
    srcItemIDCol = FindColumn(sourceSheet, "Item ID")
    srcAbbrevCol = FindColumn(sourceSheet, "Abbreviation")
    srcNameCol = FindColumn(sourceSheet, "Name")
    srcRespCol = FindColumn(sourceSheet, "Responsible")
    srcVersionCol = FindColumn(sourceSheet, "Version")
    
    ' Validate all columns found
    If srcIdentDateCol = 0 Or srcLocationCol = 0 Or srcItemIDCol = 0 Or _
       srcAbbrevCol = 0 Or srcNameCol = 0 Or srcRespCol = 0 Or srcVersionCol = 0 Then
        MsgBox "Could not find all required columns in Technical Data sheet." & vbCrLf & _
               "Required: Identified Date, Location, Item ID, Abbreviation, Name, Responsible, Version", _
               vbCritical
        Exit Sub
    End If
    
    ' Extract data from source row
    identDate = sourceSheet.Cells(sourceRow, srcIdentDateCol).Value
    location = sourceSheet.Cells(sourceRow, srcLocationCol).Value
    itemID = sourceSheet.Cells(sourceRow, srcItemIDCol).Value
    abbreviation = sourceSheet.Cells(sourceRow, srcAbbrevCol).Value
    name = sourceSheet.Cells(sourceRow, srcNameCol).Value
    responsible = sourceSheet.Cells(sourceRow, srcRespCol).Value
    version = sourceSheet.Cells(sourceRow, srcVersionCol).Value
    
    ' Validate Item ID exists
    If IsEmpty(itemID) Or Trim(CStr(itemID)) = "" Then
        MsgBox "Item ID is empty in selected row.", vbExclamation
        Exit Sub
    End If
    
    sourceItemID = CStr(itemID)
    
    ' Find column indices in target sheet (Technical File)
    Dim tgtIdentDateCol As Long, tgtLocationCol As Long, tgtItemIDCol As Long
    Dim tgtAbbrevCol As Long, tgtNameCol As Long, tgtRespCol As Long, tgtVersionCol As Long
    
    tgtIdentDateCol = FindColumn(targetSheet, "Identified Date")
    tgtLocationCol = FindColumn(targetSheet, "Location")
    tgtItemIDCol = FindColumn(targetSheet, "Item ID")
    tgtAbbrevCol = FindColumn(targetSheet, "Abbreviation")
    tgtNameCol = FindColumn(targetSheet, "Name")
    tgtRespCol = FindColumn(targetSheet, "Responsible")
    tgtVersionCol = FindColumn(targetSheet, "Version")
    
    ' Validate all columns found in target
    If tgtIdentDateCol = 0 Or tgtLocationCol = 0 Or tgtItemIDCol = 0 Or _
       tgtAbbrevCol = 0 Or tgtNameCol = 0 Or tgtRespCol = 0 Or tgtVersionCol = 0 Then
        MsgBox "Could not find all required columns in Technical File sheet." & vbCrLf & _
               "Required: Identified Date, Location, Item ID, Abbreviation, Name, Responsible, Version", _
               vbCritical
        Exit Sub
    End If
    
    ' Find insert position in Technical File based on Item ID ordering
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, tgtItemIDCol).End(xlUp).row
    insertPosition = lastRow + 1 ' Default to end
    
    For i = 7 To lastRow ' Assuming row 1-6 is header
        Dim targetItemID As String
        targetItemID = CStr(targetSheet.Cells(i, tgtItemIDCol).Value)
        
        ' Insert before the first Item ID that's greater than source Item ID
        If StrComp(sourceItemID, targetItemID, vbTextCompare) < 0 Then
            insertPosition = i
            Exit For
        End If
    Next i
    
    ' Activate target sheet and select the insert row
    targetSheet.Activate
    targetSheet.Rows(insertPosition).Select
    
    ' Call the matrix insert function
    ' This will insert row and column at the current selection
    Call InsertMatrixRowAndColumn
    
    ' Copy data to the newly inserted row
    ' Note: InsertMatrixRowAndColumn already sets the date, so we'll overwrite if needed
    targetSheet.Cells(insertPosition, tgtIdentDateCol).Value = identDate
    targetSheet.Cells(insertPosition, tgtLocationCol).Value = location
    targetSheet.Cells(insertPosition, tgtItemIDCol).Value = itemID
    targetSheet.Cells(insertPosition, tgtAbbrevCol).Value = abbreviation
    targetSheet.Cells(insertPosition, tgtNameCol).Value = name
    targetSheet.Cells(insertPosition, tgtRespCol).Value = responsible
    targetSheet.Cells(insertPosition, tgtVersionCol).Value = version
    
    ' Inform user
    MsgBox "Row copied to Technical File at position " & insertPosition & ".", vbInformation
    
    ' Return to source sheet
    sourceSheet.Activate
    sourceSheet.Rows(sourceRow).Select
End Sub


' Helper Function FindColumn()
' Purpose: Find column index by header name
' Parameters:
' - ws: Worksheet to search
' - headerName: Column header to find
' Returns: Column index (1-based), or 0 if not found

Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim col As Long
    Dim lastCol As Long
    
    FindColumn = 0
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If Trim(UCase(ws.Cells(3, col).Value)) = Trim(UCase(headerName)) Then
            FindColumn = col
            Exit Function
        End If
    Next col
End Function


' ========================================
' Workbook Event Handlers
' ========================================
' Add these to ThisWorkbook module to auto-add/remove context menu items

' Private Sub Workbook_Open()
'     Call AddCopyToTechFileContextMenu
'     Call AddMatrixInsertToRowContextMenu
' End Sub

' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     Call RemoveCopyToTechFileContextMenu
'     Call RemoveMatrixInsertFromRowContextMenu
' End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'     ' Optional: Only show menu on specific sheets
     If Sh.name = "Technical Data" Then
         Call AddCopyToTechFileContextMenu
     Else
         Call RemoveCopyToTechFileContextMenu
     End If
End Sub
