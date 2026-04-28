Option Explicit

' =============================================================
' STEP 1 – Run SetupImportButton() once to create the
'          "Combined TBs & Analysis" sheet and the Import TB
'          button.  After that, use the button to run imports.
' =============================================================

Sub SetupImportButton()
    Dim ws As Worksheet
    Set ws = GetOrCreateCombinedSheet(ThisWorkbook)
    AddImportButton ws
    ws.Activate
    MsgBox "Setup complete!" & vbNewLine & _
           "Use the 'Import TB' button on the '" & ws.Name & "' sheet.", _
           vbInformation, "Setup"
End Sub


' -------------------------------------------------------
' Creates "Combined TBs & Analysis" if not already present.
' Returns the sheet either way.
' -------------------------------------------------------
Function GetOrCreateCombinedSheet(wb As Workbook) As Worksheet
    Const SHEET_NAME As String = "Combined TBs & Analysis"
    Dim ws As Worksheet

    For Each ws In wb.Sheets
        If ws.Name = SHEET_NAME Then
            Set GetOrCreateCombinedSheet = ws
            Exit Function
        End If
    Next ws

    ' --- Create sheet ---
    Application.ScreenUpdating = False
    Set ws = wb.Sheets.Add(Before:=wb.Sheets(1))
    ws.Name = SHEET_NAME

    ' Row 1: dedicated button row
    ws.Rows(1).RowHeight = 35

    ' Row 2: column headers
    Dim headers As Variant
    headers = Array("Project", "Entity", "Account", "Closing Balance")
    Dim col As Integer
    For col = 1 To 4
        ws.Cells(2, col).Value = headers(col - 1)
    Next col

    ' Style header row (dark navy, white bold text)
    With ws.Range("A2:D2")
        .Font.Bold           = True
        .Font.Color          = RGB(255, 255, 255)
        .Font.Name           = "Calibri"
        .Interior.Color      = RGB(0, 32, 96)
        .HorizontalAlignment = xlLeft
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight    = xlMedium
            .Color     = RGB(255, 255, 255)
        End With
    End With

    ' AutoFilter on header row
    ws.Range("A2:D2").AutoFilter

    ' Freeze rows 1-2 (button + headers always visible)
    Application.ScreenUpdating = True
    ws.Activate
    ws.Rows(3).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Set GetOrCreateCombinedSheet = ws
End Function


' -------------------------------------------------------
' Adds (or replaces) the Import TB button in row 1.
' Can also be run standalone from Alt+F8 to recreate it.
' -------------------------------------------------------
Sub AddImportButton(Optional ws As Worksheet = Nothing)
    Const SHEET_NAME As String = "Combined TBs & Analysis"

    If ws Is Nothing Then
        Dim wsFind As Worksheet
        For Each wsFind In ThisWorkbook.Sheets
            If wsFind.Name = SHEET_NAME Then
                Set ws = wsFind
                Exit For
            End If
        Next wsFind
        If ws Is Nothing Then
            MsgBox "Sheet '" & SHEET_NAME & "' not found." & vbNewLine & _
                   "Run SetupImportButton() first.", vbExclamation
            Exit Sub
        End If
    End If

    ' Remove existing button if present
    Dim btn As Object
    For Each btn In ws.Buttons
        If btn.Name = "btnImportTB" Then
            btn.Delete
            Exit For
        End If
    Next btn

    ' Place button in row 1, anchored to column A
    Dim newBtn As Object
    Set newBtn = ws.Buttons.Add( _
        ws.Cells(1, 1).Left + 4, _
        ws.Cells(1, 1).Top  + 4, _
        120, _
        ws.Rows(1).RowHeight - 8)
    With newBtn
        .Caption   = "Import TB"
        .OnAction  = "ImportTB"
        .Name      = "btnImportTB"
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub
