Option Explicit

' =============================================================
' Trial Balance Importer
' Usage: click the "Import TB" button on the Combined sheet,
'        or run ImportTB() from the Macro menu (Alt+F8)
'
' Combined sheet layout:
'   Row 1  : Import TB button (dedicated button row)
'   Row 2  : Headers  (Project | Entity | Account | Closing Balance)
'   Row 3+ : Data
' =============================================================

' -------------------------------------------------------
' MAIN ENTRY POINT
' -------------------------------------------------------
Sub ImportTB()
    Dim wb              As Workbook
    Dim srcWb           As Workbook
    Dim srcWs           As Worksheet
    Dim combinedWs      As Worksheet
    Dim newWs           As Worksheet
    Dim anchorWs        As Worksheet
    Dim lastInsertedWs  As Worksheet
    Dim fd              As FileDialog
    Dim filePath        As String
    Dim tbName          As String
    Dim dataStartRow    As Long
    Dim sharedDataStart As Long
    Dim lastRow         As Long
    Dim firstNewRow     As Long
    Dim nextRow         As Long
    Dim clearLast       As Long
    Dim filesOK         As Long
    Dim i               As Long
    Dim j               As Integer
    Dim s               As Integer
    Dim sheetIndex      As Integer
    Dim sameFormat      As Boolean
    Dim hInput          As String
    Dim dInput          As String
    Dim sheetList       As String
    Dim sheetChoice     As String

    Set wb = ThisWorkbook

    ' Create (or retrieve) the Combined sheet
    Set combinedWs = GetOrCreateCombinedSheet(wb)

    ' Resolve insertion anchor – new TB sheets go after "Supportings >>"
    Const ANCHOR_SHEET As String = "Supportings >>"
    On Error Resume Next
    Set anchorWs = wb.Sheets(ANCHOR_SHEET)
    On Error GoTo 0
    If anchorWs Is Nothing Then
        MsgBox "Sheet '" & ANCHOR_SHEET & "' not found." & vbNewLine & _
               "TB sheets will be inserted at the end of the workbook.", vbExclamation
        Set anchorWs = wb.Sheets(wb.Sheets.Count)
    End If
    Set lastInsertedWs = anchorWs

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

    ' ---------- Same or different format? ----------
    If fd.SelectedItems.Count > 1 Then
        sameFormat = (MsgBox( _
            "You selected " & fd.SelectedItems.Count & " files." & vbNewLine & vbNewLine & _
            "Do all files have the SAME format?" & vbNewLine & _
            "(same header row and data start row)" & vbNewLine & vbNewLine & _
            "Yes = ask once for all files" & vbNewLine & _
            "No  = ask separately for each file", _
            vbYesNo + vbQuestion, "File Format") = vbYes)
    Else
        sameFormat = True
    End If

    ' If same format, ask once up front
    If sameFormat Then
        hInput = InputBox("Enter the ROW NUMBER that contains the column headers.", _
                          "Header Row (all files)", "1")
        If hInput = "" Or Not IsNumeric(hInput) Then
            MsgBox "Import cancelled.", vbInformation
            Exit Sub
        End If
        dInput = InputBox("Header is on row " & hInput & "." & vbNewLine & _
                          "Enter the ROW NUMBER where data starts.", _
                          "Data Start Row (all files)", CStr(CLng(hInput) + 1))
        If dInput = "" Or Not IsNumeric(dInput) Then
            MsgBox "Import cancelled.", vbInformation
            Exit Sub
        End If
        sharedDataStart = CLng(dInput)
    End If

    filesOK = 0

    ' Clear existing data (rows 3+) and start writing from row 3
    Application.ScreenUpdating = False
    clearLast = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row
    If clearLast > 2 Then combinedWs.Rows("3:" & clearLast).Delete Shift:=xlUp
    Application.ScreenUpdating = True
    nextRow = 3

    For j = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(j)
        tbName   = GetFileNameWithoutExt(filePath)

        ' ======================================================
        ' STEP 1 – Open source workbook and copy its sheet
        ' ======================================================
        Application.ScreenUpdating = False
        Application.DisplayAlerts  = False

        On Error GoTo OpenError
        Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0

        ' If source has multiple sheets, ask which one to use
        sheetIndex = 1
        If srcWb.Sheets.Count > 1 Then
            Application.ScreenUpdating = True
            sheetList = ""
            For s = 1 To srcWb.Sheets.Count
                sheetList = sheetList & s & " - " & srcWb.Sheets(s).Name & vbNewLine
            Next s
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
        Set lastInsertedWs = newWs

        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True

        ' ======================================================
        ' STEP 2 – Determine data start row
        ' ======================================================
        If sameFormat Then
            dataStartRow = sharedDataStart
        Else
            hInput = InputBox("File: " & tbName & vbNewLine & vbNewLine & _
                              "Enter the ROW NUMBER that contains the column headers.", _
                              "Header Row", "1")
            If hInput = "" Or Not IsNumeric(hInput) Then GoTo SkipFile

            dInput = InputBox("File: " & tbName & vbNewLine & _
                              "Header is on row " & hInput & "." & vbNewLine & _
                              "Enter the ROW NUMBER where data starts.", _
                              "Data Start Row", CStr(CLng(hInput) + 1))
            If dInput = "" Or Not IsNumeric(dInput) Then GoTo SkipFile

            dataStartRow = CLng(dInput)
        End If

        ' ======================================================
        ' STEP 3 – Copy rows to Combined sheet
        '
        ' Mapping (source -> combined):
        '   n/a literal  -> col A  Project
        '   TB file name -> col B  Entity
        '   source col B -> col C  Account
        '   source col F -> col D  Closing Balance
        ' ======================================================
        lastRow = newWs.Cells(newWs.Rows.Count, "B").End(xlUp).Row

        If lastRow < dataStartRow Then
            MsgBox "No data found in column B from row " & dataStartRow & _
                   " in '" & tbName & "'. Skipping.", vbExclamation
            GoTo SkipFile
        End If

        Application.ScreenUpdating = False
        firstNewRow = nextRow

        For i = dataStartRow To lastRow
            If Trim(CStr(newWs.Cells(i, 2).Value)) = "" Then GoTo NextRow

            combinedWs.Cells(nextRow, 1).Value = "n/a"
            combinedWs.Cells(nextRow, 2).Value = tbName
            combinedWs.Cells(nextRow, 3).Value = newWs.Cells(i, 2).Value  ' src col B
            combinedWs.Cells(nextRow, 4).Value = newWs.Cells(i, 6).Value  ' src col F

            nextRow = nextRow + 1
NextRow:
        Next i

        ' Number formatting for closing balance
        If nextRow > firstNewRow Then
            combinedWs.Range( _
                combinedWs.Cells(firstNewRow, 4), _
                combinedWs.Cells(nextRow - 1, 4)) _
                .NumberFormat = "#,##0.00;[Red]-#,##0.00"
        End If

        filesOK = filesOK + 1
        Application.ScreenUpdating = True

SkipFile:
    Next j

    ' Final tidy-up
    Application.ScreenUpdating = False
    combinedWs.Columns("A:D").AutoFit
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

    ' --- Row 1: dedicated button row ---
    ws.Rows(1).RowHeight = 35

    ' --- Row 2: column headers ---
    Dim headers As Variant
    headers = Array("Project", "Entity", "Account", "Closing Balance")
    Dim col As Integer
    For col = 1 To 4
        ws.Cells(2, col).Value = headers(col - 1)
    Next col

    ' --- Style header row (dark navy, white bold text) ---
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

    ' --- AutoFilter on header row ---
    ws.Range("A2:D2").AutoFilter

    ' --- Freeze rows 1-2 so button and headers stay visible ---
    Application.ScreenUpdating = True
    ws.Activate
    ws.Rows(3).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    ' --- Add Import TB button in row 1 ---
    AddImportButton ws

    Set GetOrCreateCombinedSheet = ws
End Function


' -------------------------------------------------------
' Adds (or replaces) the Import TB button in row 1.
' Run this manually (Alt+F8 > AddImportButton) to
' recreate the button if it was ever deleted.
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
            MsgBox "Sheet '" & SHEET_NAME & "' not found.", vbExclamation
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

    ' Place button in row 1, starting at column A
    Dim btnLeft   As Double: btnLeft   = ws.Cells(1, 1).Left + 4
    Dim btnTop    As Double: btnTop    = ws.Cells(1, 1).Top  + 4
    Dim btnWidth  As Double: btnWidth  = 120
    Dim btnHeight As Double: btnHeight = ws.Rows(1).RowHeight - 8

    Dim newBtn As Object
    Set newBtn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    With newBtn
        .Caption   = "Import TB"
        .OnAction  = "ImportTB"
        .Name      = "btnImportTB"
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub


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
