Attribute VB_Name = "RowColumnModule"
'Initialization

' Function AddMatrixInsertToRowContextMenu()
' Purpose: Add insert option to context Menu
' Key actions:
' - Add control
' Code flow Summary:
' - Obtains Context Menu, if fails informs
' - Add control, if fails informs
' - Configures Control

Sub AddMatrixInsertToRowContextMenu()
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
        .Caption = "Insert Row-Column"
        .OnAction = "InsertMatrixRowAndColumn"
        .BeginGroup = True
    End With
End Sub


' Function RemoveMatrixInsertFromRowContextMenu()
' Purpose: Remove insert option from context Menu
' Key actions:
' - Remove control
' Code flow Summary:
' - Obtains Context Menu, if fails exit
' - Find Row Coolumn menu
' - Remove menu
Sub RemoveMatrixInsertFromRowContextMenu()
    Dim ctrl As CommandBarControl
    Dim contextMenu As CommandBar

    Set contextMenu = Application.CommandBars("Row")
    If contextMenu Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ctrl In contextMenu.Controls
        If ctrl.Caption = "Insert Row-Column" Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub

' Function InsertMatrixRowAndColumn()
' Purpose: Main function that inserts a column on the N_squared relationship matrix when a row is inserted.
'          Secondary function, adds date to inserted row.
' Key actions:
' - Insert Column
' - Applies correct format to cell (Color)
' - Set values to specific cells: Name of column and date to row.
' Code flow Summary:
' - load named ranges from Namespace and check existance
' - Find where to insert
' - Insert
' - Format Relationship cell: White around and gray center
' - Set Values: Relationship name based on type column and date
' Important: Type column is hard coded.

Sub InsertMatrixRowAndColumn()
    Dim matrixStart As Range, matrixEnd As Range
    Dim matrixSize As Long
    Dim insertIndex As Long
    Dim identStart As Range, relStart As Range
    Dim targetCell As Range
    Dim identValue As Variant
    Dim relRowIndex As Long
    Dim formatingCell As Range

    ' Load named ranges
    On Error Resume Next
    Set matrixStart = Range("MatrixTopLeft")
    Set matrixEnd = Range("MatrixBottomRight")
    Set identStart = Range("IDENT_START")
    Set relStart = Range("REL_START")
    On Error GoTo 0
    ' Check if names exist, if not exit
    If matrixStart Is Nothing Or matrixEnd Is Nothing Then
        MsgBox "MatrixTopLeft or MatrixBottomRight not defined.", vbCritical
        Exit Sub
    End If
    If identStart Is Nothing Then
        MsgBox "IDENT_START not defined.", vbCritical
        Exit Sub
    End If
    If relStart Is Nothing Then
        MsgBox "REL_START not defined.", vbCritical
        Exit Sub
    End If

    matrixSize = Application.WorksheetFunction.Max( _
        matrixEnd.row - matrixStart.row + 1, _
        matrixEnd.Column - matrixStart.Column + 1)
    ' Find where to insert
    insertIndex = Selection.row - matrixStart.row + 1
    If insertIndex < 1 Or insertIndex > matrixSize + 1 Then
        MsgBox "Please right-click a row within the matrix.", vbExclamation
        Exit Sub
    End If

    ' Insert row and column
    matrixStart.Offset(insertIndex - 1, 0).Resize(1, matrixSize).EntireRow.Insert
    matrixStart.Offset(0, insertIndex - 1).Resize(matrixSize, 1).EntireColumn.Insert

    ' Format Relationship Cell:
    ' (1) Clear format around
    Set formatingCell = matrixStart.Offset(insertIndex - 2, insertIndex - 1)
    formatingCell.Interior.Color = RGB(255, 255, 255) ' White
    Set formatingCell = matrixStart.Offset(insertIndex - 1, insertIndex - 2)
    formatingCell.Interior.Color = RGB(255, 255, 255) ' White
    ' (2) Highlight the inserted row's new column cell in light gray
    Set formatingCell = matrixStart.Offset(insertIndex - 1, insertIndex - 1)
    formatingCell.Interior.Color = RGB(128, 128, 128) ' gray
    
    
    ' Set automatic values:
    ' (1) Type (referenced) for the column
    ' IMPORTANT: IT looks for a fixed value of 4 from the Identification column.
    identValue = "=" & Cells(matrixStart.row + insertIndex - 1, identStart.Column + 4).Address(External:=True)
    relRowIndex = relStart.row + 1
    Set targetCell = Cells(relRowIndex, matrixStart.Column + insertIndex - 1)
    targetCell.Value = identValue
    ' (2) Date on first column
    Set targetCell = Cells(matrixStart.row + insertIndex - 1, identStart.Column)
    targetCell.Value = Format(Now, "yyyy.mm.dd")

    ' Select inserted row
    matrixStart.Offset(insertIndex - 1, 0).Select
End Sub


