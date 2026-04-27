Option Explicit

' =============================================================
' Trial Balance Importer
' Usage: Run ImportTB() from the Macro menu
' =============================================================

' -------------------------------------------------------
' MAIN ENTRY POINT
' -------------------------------------------------------
Sub ImportTB()
    Dim wb          As Workbook
    Dim srcWb       As Workbook
    Dim srcWs       As Worksheet
    Dim combinedWs  As Worksheet
    Dim newWs       As Worksheet
    Dim fd          As FileDialog
    Dim filePath    As String
    Dim tbName      As String
    Dim dataStartRow As Long
    Dim lastRow     As Long
    Dim i           As Long
    Dim nextRow     As Long
    Dim filesOK         As Long
    Dim j               As Integer
    Dim lastInsertedWs  As Worksheet   ' tracks insert position after "Supportings >>"

    Set wb = ThisWorkbook

    ' Create (or retrieve) the Combined sheet BEFORE touching anything else
    Set combinedWs = GetOrCreateCombinedSheet(wb)

    ' Resolve insertion anchor – new TB sheets go after "Supportings >>"
    Const ANCHOR_SHEET As String = "Supportings >>"
    Dim anchorWs As Worksheet
    On Error Resume Next
    Set anchorWs = wb.Sheets(ANCHOR_SHEET)
    On Error GoTo 0
    If anchorWs Is Nothing Then
        MsgBox "Sheet '" & ANCHOR_SHEET & "' not found." & vbNewLine & _
               "TB sheets will be inserted at the end of the workbook.", vbExclamation
        Set anchorWs = wb.Sheets(wb.Sheets.Count)
    End If
    Set lastInsertedWs = anchorWs   ' first TB goes right after the anchor

    ' ---------- File picker ----------
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = True
        .Title = "Select Trial Balance File(s) to Import"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
    End With

    If fd.Show = False Then
        MsgBox "No files selected. Import cancelled.", vbInformation
        Exit Sub
    End If

    filesOK = 0

    For j = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(j)
        tbName   = GetFileNameWithoutExt(filePath)

        ' ======================================================
        ' STEP 1 – Open source workbook and copy its first sheet
        '          as a new tab in this workbook
        ' ======================================================
        Application.ScreenUpdating = False
        Application.DisplayAlerts  = False

        On Error GoTo OpenError
        Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0

        ' If source has multiple sheets, ask which one to use
        Dim sheetIndex As Integer
        sheetIndex = 1
        If srcWb.Sheets.Count > 1 Then
            Application.ScreenUpdating = True
            Dim sheetList As String
            Dim s As Integer
            sheetList = ""
            For s = 1 To srcWb.Sheets.Count
                sheetList = sheetList & s & " - " & srcWb.Sheets(s).Name & vbNewLine
            Next s
            Dim sheetChoice As String
            sheetChoice = InputBox( _
                "File: " & tbName & " has " & srcWb.Sheets.Count & " sheets:" & _
                vbNewLine & sheetList & vbNewLine & _
                "Enter the sheet NUMBER to import:", "Select Sheet", "1")
            If sheetChoice = "" Or Not IsNumeric(sheetChoice) Then
                srcWb.Close SaveChanges:=False
                GoTo SkipFile
            End If
            sheetIndex = CInt(sheetChoice)
            If sheetIndex < 1 Or sheetIndex > srcWb.Sheets.Count Then sheetIndex = 1
            Application.ScreenUpdating = False
        End If

        Set srcWs = srcWb.Sheets(sheetIndex)

        ' Copy entire sheet to this workbook (preserves all formatting)
        srcWs.Copy After:=lastInsertedWs
        Set newWs          = wb.Sheets(lastInsertedWs.Index + 1)
        newWs.Name         = GetUniqueSheetName(wb, tbName)
        Set lastInsertedWs = newWs   ' next TB inserts after this one

        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True

        dataStartRow = 7

        ' ======================================================
        ' STEP 2 – Copy rows to Combined sheet
        '
        ' Mapping (source -> combined):
        '   n/a literal          -> col A  Project
        '   TB file name         -> col B  Entity
        '   source col B         -> col C  Account
        '   col D left blank     ->        (user writes here)
        '   source col D         -> col E  Transaction type
        '   (col F intentionally left blank for user use)
        '   source col F         -> col G  Closing Balance
        ' ======================================================
        lastRow = newWs.Cells(newWs.Rows.Count, "B").End(xlUp).Row

        If lastRow < dataStartRow Then
            MsgBox "No data found in column B from row " & dataStartRow & _
                   " in '" & tbName & "'. Skipping.", vbExclamation
            GoTo SkipFile
        End If

        Application.ScreenUpdating = False
        nextRow = GetNextDataRow(combinedWs)

        For i = dataStartRow To lastRow
            ' Skip blank Account rows
            If Trim(CStr(newWs.Cells(i, 2).Value)) = "" Then GoTo NextRow

            combinedWs.Cells(nextRow, 1).Value = "n/a"                     ' Project
            combinedWs.Cells(nextRow, 2).Value = tbName                    ' Entity        (chosen file name)
            combinedWs.Cells(nextRow, 3).Value = newWs.Cells(i, 2).Value  ' Account       (src col B)
            ' col D: left blank – user writes here
            combinedWs.Cells(nextRow, 5).Value = newWs.Cells(i, 4).Value  ' Txn type      (src col D)
            ' col F: intentionally left blank
            combinedWs.Cells(nextRow, 7).Value = newWs.Cells(i, 6).Value  ' Closing bal   (src col F)

            nextRow = nextRow + 1
NextRow:
        Next i

        ' Light number formatting for closing balance column
        If nextRow > GetNextDataRow(combinedWs) Then
            combinedWs.Range( _
                combinedWs.Cells(GetNextDataRow(combinedWs), 7), _
                combinedWs.Cells(nextRow - 1, 7)) _
                .NumberFormat = "#,##0.00;[Red]-#,##0.00"
        End If

        filesOK = filesOK + 1
        Application.ScreenUpdating = True

SkipFile:
    Next j

    ' Final tidy-up
    Application.ScreenUpdating = False
    combinedWs.Columns("A:G").AutoFit
    combinedWs.Range("A1").Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    combinedWs.Activate
    MsgBox filesOK & " of " & fd.SelectedItems.Count & _
           " Trial Balance(s) imported successfully!", vbInformation, "Import Complete"
    Exit Sub

OpenError:
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    MsgBox "Could not open file:" & vbNewLine & filePath & _
           vbNewLine & vbNewLine & Err.Description, vbCritical, "File Open Error"
    On Error GoTo 0
    Resume SkipFile
End Sub


' -------------------------------------------------------
' Creates "Combined TBs & Analysis" if not already present
' -------------------------------------------------------
Function GetOrCreateCombinedSheet(wb As Workbook) As Worksheet
    Const SHEET_NAME As String = "Combined TBs & Analysis"
    Dim ws  As Worksheet

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

    ' --- Write headers in row 1 ---
    Dim headers As Variant
    headers = Array("Project", "Entity", "Account", "", _
                    "Transaction type", "", "Closing Balance")
    Dim col As Integer
    For col = 1 To 7
        ws.Cells(1, col).Value = headers(col - 1)
    Next col

    ' --- Style header row (dark navy, white bold text) ---
    With ws.Range("A1:G1")
        .Font.Bold              = True
        .Font.Color             = RGB(255, 255, 255)
        .Font.Name              = "Calibri"
        .Interior.Color         = RGB(0, 32, 96)
        .HorizontalAlignment    = xlLeft
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight    = xlMedium
            .Color     = RGB(255, 255, 255)
        End With
    End With

    ' Add AutoFilter
    ws.Range("A1:G1").AutoFilter

    ' Freeze header row
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Application.ScreenUpdating = True
    Set GetOrCreateCombinedSheet = ws
End Function


' -------------------------------------------------------
' Extracts filename (no extension) from a full file path
' -------------------------------------------------------
Function GetFileNameWithoutExt(fullPath As String) As String
    Dim name As String
    name = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(name, ".") > 0 Then
        name = Left(name, InStrRev(name, ".") - 1)
    End If
    GetFileNameWithoutExt = name
End Function


' -------------------------------------------------------
' Ensures sheet name is unique and <= 31 characters
' -------------------------------------------------------
Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim candidate As String
    Dim suffix    As Integer
    Dim ws        As Worksheet
    Dim found     As Boolean

    If Len(baseName) > 31 Then baseName = Left(baseName, 31)
    candidate = baseName
    suffix    = 1

    Do
        found = False
        For Each ws In wb.Sheets
            If LCase(ws.Name) = LCase(candidate) Then
                found = True
                Exit For
            End If
        Next ws
        If Not found Then Exit Do
        candidate = Left(baseName, 28) & "_" & suffix
        suffix = suffix + 1
    Loop

    GetUniqueSheetName = candidate
End Function


' -------------------------------------------------------
' Returns the first empty row below existing data
' -------------------------------------------------------
Function GetNextDataRow(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(1, 1).Value = "" Then
        GetNextDataRow = 1
    Else
        GetNextDataRow = lastRow + 1
    End If
End Function


' -------------------------------------------------------
' Optional utility: clear all data rows in Combined sheet
' (keeps the header row and all TB sheets intact)
' -------------------------------------------------------
Sub ClearCombinedData()
    Const SHEET_NAME As String = "Combined TBs & Analysis"
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SHEET_NAME Then
            If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
                Dim lastRow As Long
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                ws.Rows("2:" & lastRow).Delete Shift:=xlUp
                MsgBox "Combined data cleared (header kept).", vbInformation
            Else
                MsgBox "Nothing to clear – Combined sheet already empty.", vbInformation
            End If
            Exit Sub
        End If
    Next ws
    MsgBox "Sheet '" & SHEET_NAME & "' not found.", vbExclamation
End Sub
