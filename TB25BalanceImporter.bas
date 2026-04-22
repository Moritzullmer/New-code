Attribute VB_Name = "TB25BalanceImporter"
Option Explicit

' =============================================================
' CONSTANTS
' Edit these if the sheet layout or TB file structure changes.
' =============================================================
Private Const CONSO_SHEET       As String = "Conso BS 25"
Private Const ROW_ENTITY_NAME   As Long = 15   ' entity display name in this row (= folder name on disk)
Private Const ROW_DATA_START    As Long = 17   ' first row to scan for account numbers in col A
Private Const COL_ACCOUNT       As Long = 1    ' column A holds account numbers
Private Const TB_COL_ACCT       As Long = 1    ' TB25 file: account number in col A
Private Const TB_COL_SALDO      As Long = 6    ' TB25 file: ending balance (Saldo) in col F
Private Const TB_ROW_DATA_START As Long = 2    ' TB25 file: header in row 1, data from row 2
Private Const TB_SUBFOLDER      As String = "TB 25"

' =============================================================
' MODULE-LEVEL STATE
' =============================================================
Private m_SkipLog As String

' =============================================================
' UDT — one instance per entity found in row 15
' =============================================================
Private Type EntityInfo
    EntityName  As String   ' raw text from row 15 (= folder name on disk)
    FirstCol    As Long     ' leftmost column of entity span in row 15
    LastCol     As Long     ' rightmost column of entity span
    InsertCol   As Long     ' column where "as per TB" will be inserted (= LastCol + 1)
    TBSheetName As String   ' destination sheet name in this workbook ("TB_" + truncated name)
    Skipped     As Boolean  ' True if TB file was not found or failed to import
End Type

' =============================================================
' PUBLIC ENTRY POINT
' Prompts for the parent folder containing entity sub-folders,
' imports each entity's TB25 file as a new sheet, then inserts
' an "as per TB" column to the right of each entity column in
' the "Conso BS 25" sheet and fills it with account sums.
' =============================================================
Public Sub ImportTB25Balances()

    Dim fd             As FileDialog
    Dim parentFolder   As String
    Dim wsBS           As Worksheet
    Dim entities()     As EntityInfo
    Dim entCount       As Long
    Dim processedCount As Long
    Dim i              As Long
    Dim tbPath         As String
    Dim wsTB           As Worksheet
    Dim errMsg         As String

    m_SkipLog = ""  ' clear log from any previous run

    ' ---- Select parent folder ----
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select the parent folder containing entity sub-folders"
    If fd.Show <> -1 Then Exit Sub
    parentFolder = fd.SelectedItems(1)

    ' ---- Validate "Conso BS 25" sheet ----
    On Error Resume Next
    Set wsBS = ThisWorkbook.Sheets(CONSO_SHEET)
    On Error GoTo 0
    If wsBS Is Nothing Then
        MsgBox "Sheet '" & CONSO_SHEET & "' not found in this workbook.", _
               vbExclamation, "Sheet Missing"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    AppPerfOn

    ' ---- Discover entities from row 15 ----
    FindEntityColumns wsBS, entities, entCount

    If entCount = 0 Then
        MsgBox "No entity names found in row " & ROW_ENTITY_NAME & " of '" & CONSO_SHEET & "'.", _
               vbInformation, "Nothing to Do"
        GoTo CleanExit
    End If

    ' ---- Phase 1: Import TB25 files (left to right, before any column insertion) ----
    For i = 0 To entCount - 1
        tbPath = FindTB25File(parentFolder, entities(i).EntityName)
        If tbPath = "" Then
            LogSkip "Entity '" & entities(i).EntityName & "': TB25 file not found under '" & _
                    parentFolder & "\" & entities(i).EntityName & "\" & TB_SUBFOLDER & "' — skipped."
            entities(i).Skipped = True
        Else
            If Not ImportSheetData(tbPath, entities(i).TBSheetName) Then
                LogSkip "Entity '" & entities(i).EntityName & "': could not open '" & tbPath & "' — skipped."
                entities(i).Skipped = True
            Else
                processedCount = processedCount + 1
            End If
        End If
    Next i

    ' ---- Phase 2: Insert "as per TB" columns (RIGHT TO LEFT to prevent index drift) ----
    For i = entCount - 1 To 0 Step -1
        If Not entities(i).Skipped Then
            Set wsTB = Nothing
            On Error Resume Next
            Set wsTB = ThisWorkbook.Sheets(entities(i).TBSheetName)
            On Error GoTo ErrHandler
            If Not wsTB Is Nothing Then
                InsertAndFillColumn wsBS, entities(i), wsTB
            End If
        End If
    Next i

CleanExit:
    AppPerfOff
    If m_SkipLog <> "" Then
        MsgBox "Import complete (" & processedCount & " entit(y/ies) processed) with warnings:" & _
               vbCrLf & vbCrLf & m_SkipLog, vbExclamation, "TB25 Import — Warnings"
    Else
        MsgBox "TB25 import complete. " & processedCount & " entit(y/ies) processed.", _
               vbInformation, "TB25 Import — Done"
    End If
    Exit Sub

ErrHandler:
    errMsg = Err.Description
    AppPerfOff
    MsgBox "Unexpected error:" & vbCrLf & errMsg, vbExclamation, "TB25 Import Error"

End Sub

' =============================================================
' PHASE SUBROUTINES
' =============================================================

' InsertAndFillColumn
' Inserts a new column at ent.InsertCol (handling idempotency),
' writes "as per TB" header in ROW_ENTITY_NAME, then fills each
' account-bearing row with the sum from the TB25 sheet.
Private Sub InsertAndFillColumn(wsBS As Worksheet, ent As EntityInfo, wsTB As Worksheet)

    Dim insertAt    As Long
    Dim lastDataRow As Long
    Dim r           As Long
    Dim acctVal     As String

    insertAt = ent.InsertCol

    ' Idempotency: if the cell at (ROW_ENTITY_NAME, insertAt) already says
    ' "as per TB" this column was written in a prior run — delete it first.
    If Trim(CStr(wsBS.Cells(ROW_ENTITY_NAME, insertAt).Value)) = "as per TB" Then
        Application.DisplayAlerts = False
        wsBS.Columns(insertAt).Delete Shift:=xlToLeft
        Application.DisplayAlerts = True
    End If

    ' Insert a blank column, pushing existing content to the right.
    wsBS.Columns(insertAt).Insert Shift:=xlToRight

    ' Write header.
    wsBS.Cells(ROW_ENTITY_NAME, insertAt).Value = "as per TB"

    ' Fill data rows.
    lastDataRow = LastUsedRow(wsBS, COL_ACCOUNT)
    For r = ROW_DATA_START To lastDataRow
        acctVal = Trim(CStr(wsBS.Cells(r, COL_ACCOUNT).Value))
        If IsAccountRow(acctVal) Then
            wsBS.Cells(r, insertAt).Value = SumAccountInTB(wsTB, acctVal)
        End If
    Next r

End Sub

' =============================================================
' DISCOVERY FUNCTIONS
' =============================================================

' FindEntityColumns
' Scans row ROW_ENTITY_NAME from col 2 onwards.
' Uses MergeArea.Columns.Count to detect merged cell spans so that
' a multi-column entity header is treated as a single entity.
' Populates the entities() array and entCount.
Private Sub FindEntityColumns(wsBS As Worksheet, _
                               ByRef entities() As EntityInfo, _
                               ByRef entCount As Long)

    Dim lastCol  As Long
    Dim c        As Long
    Dim cellVal  As String
    Dim spanCols As Long

    entCount = 0
    ReDim entities(0)

    lastCol = wsBS.Cells(ROW_ENTITY_NAME, wsBS.Columns.Count).End(xlToLeft).Column

    c = 2   ' col A is account numbers — start scanning from col B
    Do While c <= lastCol
        cellVal = Trim(CStr(wsBS.Cells(ROW_ENTITY_NAME, c).Value))
        If cellVal <> "" Then
            spanCols = wsBS.Cells(ROW_ENTITY_NAME, c).MergeArea.Columns.Count

            ReDim Preserve entities(entCount)
            entities(entCount).EntityName  = cellVal
            entities(entCount).FirstCol    = c
            entities(entCount).LastCol     = c + spanCols - 1
            entities(entCount).InsertCol   = c + spanCols   ' immediately after entity span
            entities(entCount).TBSheetName = SafeSheetName(cellVal)
            entities(entCount).Skipped     = False
            entCount = entCount + 1

            c = c + spanCols   ' jump past the entire merged span
        Else
            c = c + 1
        End If
    Loop

End Sub

' FindTB25File
' Returns the full path to the first .xlsx found under:
'   parentFolder\entityName\TB 25\           (direct)
'   parentFolder\entityName\TB 25\<sub>\     (one level deep)
' Returns "" if nothing is found.
' Uses a two-pass approach to avoid Dir() re-entrancy issues.
Private Function FindTB25File(parentFolder As String, entityName As String) As String

    Dim tb25Root  As String
    Dim f         As String
    Dim subDirs() As String
    Dim subCount  As Long
    Dim sd        As String
    Dim sdFull    As String
    Dim sdAttr    As Long
    Dim j         As Long
    Dim subFull   As String
    Dim g         As String
    Dim chk       As String

    ' Any Dir/GetAttr error (bad path chars, permissions, etc.) → treat as not found.
    On Error GoTo NotFound

    tb25Root = parentFolder & "\" & entityName & "\" & TB_SUBFOLDER
    subCount = 0

    ' Verify the TB 25 folder exists.
    chk = Dir(tb25Root, vbDirectory)
    If chk = "" Then GoTo NotFound

    ' Pass 1a: look directly in the TB 25 folder.
    f = Dir(tb25Root & "\*.xlsx")
    If f <> "" Then
        FindTB25File = tb25Root & "\" & f
        Exit Function
    End If

    ' Pass 1b: collect sub-directory names before doing nested Dir calls.
    ' (Nested Dir calls would reset the outer Dir iterator.)
    ReDim subDirs(0)
    sd = Dir(tb25Root & "\*", vbDirectory)
    Do While sd <> ""
        If sd <> "." And sd <> ".." Then
            sdFull = tb25Root & "\" & sd
            sdAttr = 0
            On Error Resume Next
            sdAttr = GetAttr(sdFull)
            On Error GoTo NotFound
            If (sdAttr And vbDirectory) = vbDirectory Then
                ReDim Preserve subDirs(subCount)
                subDirs(subCount) = sd
                subCount = subCount + 1
            End If
        End If
        sd = Dir()
    Loop

    ' Pass 2: search each sub-directory for an xlsx file.
    For j = 0 To subCount - 1
        subFull = tb25Root & "\" & subDirs(j)
        g = Dir(subFull & "\*.xlsx")
        If g <> "" Then
            FindTB25File = subFull & "\" & g
            Exit Function
        End If
    Next j

    FindTB25File = ""
    Exit Function

NotFound:
    FindTB25File = ""

End Function

' =============================================================
' DATA FUNCTIONS
' =============================================================

' ImportSheetData
' Opens filePath read-only, copies the first sheet's used-range
' values into a new sheet named sheetName in ThisWorkbook,
' then closes the source workbook.
' Returns True on success, False on failure.
Private Function ImportSheetData(filePath As String, sheetName As String) As Boolean

    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastR As Long
    Dim lastC As Long

    ' Any unexpected error → close source if open, return False (soft skip).
    On Error GoTo ImportFailed

    ' Idempotency: delete an existing sheet with the same name.
    If SheetExists(ThisWorkbook, sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    ' Open the source workbook read-only.
    ' DisplayAlerts=False suppresses "update links?" dialogs.
    Application.DisplayAlerts = False
    Set wbSrc = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)
    Application.DisplayAlerts = True

    If wbSrc Is Nothing Then GoTo ImportFailed

    Set wsSrc = wbSrc.Sheets(1)

    ' Determine source extent.
    lastR = LastUsedRow(wsSrc, TB_COL_ACCT)
    lastC = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastC < TB_COL_SALDO Then lastC = TB_COL_SALDO  ' ensure at least col F is included

    ' Create destination sheet.
    Set wsDst = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDst.Name = sheetName

    ' Copy values only (no clipboard, no formatting dependency).
    If lastR >= 1 And lastC >= 1 Then
        wsDst.Range(wsDst.Cells(1, 1), wsDst.Cells(lastR, lastC)).Value = _
            wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastR, lastC)).Value
    End If

    ' Amber tab to match existing TB sheet convention.
    wsDst.Tab.Color = RGB(255, 192, 0)

    ' Close source without saving.
    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    ImportSheetData = True
    Exit Function

ImportFailed:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ImportSheetData = False

End Function

' SumAccountInTB
' Loads the TB sheet data into a Variant array and sums column F
' (TB_COL_SALDO) for every row where column A (TB_COL_ACCT)
' exactly matches accountNum.  Equivalent to SUMIF.
Private Function SumAccountInTB(wsTB As Worksheet, accountNum As String) As Double

    Dim lastR As Long
    Dim arr   As Variant
    Dim total As Double
    Dim r     As Long

    If wsTB Is Nothing Then SumAccountInTB = 0 : Exit Function

    lastR = LastUsedRow(wsTB, TB_COL_ACCT)
    If lastR < TB_ROW_DATA_START Then SumAccountInTB = 0 : Exit Function

    arr = wsTB.Range(wsTB.Cells(TB_ROW_DATA_START, 1), _
                     wsTB.Cells(lastR, TB_COL_SALDO)).Value

    total = 0
    For r = 1 To UBound(arr, 1)
        If Trim(CStr(arr(r, TB_COL_ACCT))) = accountNum Then
            total = total + ParseGermanNumber(arr(r, TB_COL_SALDO))
        End If
    Next r

    SumAccountInTB = total

End Function

' =============================================================
' UTILITY FUNCTIONS
' =============================================================

' IsAccountRow
' Returns True when the cell value looks like an account number:
' starts with a digit (0-9) AND contains a hyphen (e.g. "1111-2050").
Private Function IsAccountRow(cellVal As Variant) As Boolean
    Dim s As String
    s = Trim(CStr(cellVal))
    If Len(s) = 0 Then IsAccountRow = False : Exit Function
    IsAccountRow = (s Like "[0-9]*") And (InStr(s, "-") > 0)
End Function

' SafeSheetName
' Returns "TB_" + first 20 chars of entityName with illegal
' sheet-name characters replaced by underscores.
' Result is at most 23 chars, well within Excel's 31-char limit.
Private Function SafeSheetName(entityName As String) As String
    Dim s As String
    s = Left(entityName, 20)
    s = Replace(s, "/",  "_")
    s = Replace(s, "\",  "_")
    s = Replace(s, "?",  "_")
    s = Replace(s, "*",  "_")
    s = Replace(s, "[",  "_")
    s = Replace(s, "]",  "_")
    s = Replace(s, ":",  "_")
    SafeSheetName = "TB_" & s
End Function

' SheetExists
' Returns True if a sheet named sheetName already exists in wb.
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' LogSkip
' Appends a message to the module-level skip log; shown as a
' single consolidated MsgBox at the end of ImportTB25Balances.
Private Sub LogSkip(msg As String)
    If m_SkipLog = "" Then
        m_SkipLog = msg
    Else
        m_SkipLog = m_SkipLog & vbCrLf & msg
    End If
End Sub

' AppPerfOn / AppPerfOff
' Renamed from SetAppPerformance to avoid the reserved word "on"
' as a parameter name causing a compile error.
Private Sub AppPerfOn()
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual
    Application.EnableEvents  = False
End Sub

Private Sub AppPerfOff()
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Application.EnableEvents  = True
    Application.DisplayAlerts = True
End Sub

' LastUsedRow
' Returns the last row index that contains data in the given column.
Private Function LastUsedRow(ws As Worksheet, col As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ParseGermanNumber
' Converts a cell value that may be stored as a German-formatted
' string (period = thousands separator, comma = decimal separator)
' to a Double.  If the value is already a numeric VBA type it is
' returned directly without string conversion.
Private Function ParseGermanNumber(rawVal As Variant) As Double
    Dim s As String

    If IsEmpty(rawVal) Or CStr(rawVal) = "" Then
        ParseGermanNumber = 0
        Exit Function
    End If

    Select Case VarType(rawVal)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbByte, vbDecimal, vbCurrency
            ParseGermanNumber = CDbl(rawVal)
            Exit Function
    End Select

    s = CStr(rawVal)
    s = Replace(s, ".", "")    ' strip thousands separator (period)
    s = Replace(s, ",", ".")   ' convert decimal comma to dot

    If IsNumeric(s) Then
        ParseGermanNumber = CDbl(s)
    Else
        ParseGermanNumber = 0
    End If
End Function
