Attribute VB_Name = "JournalEntryFormatter"
Option Explicit

' =============================================================
' ENGAGEMENT CONSTANTS
' Edit CLIENT_NAME, WP_REF, and PERIOD_END before each engagement
' =============================================================
Private Const CLIENT_NAME As String = "[CLIENT_NAME]"
Private Const WP_REF      As String = "[WP_REF]"
Private Const PERIOD_END  As String = "[PERIOD_END]"

' =============================================================
' PUBLIC ENTRY POINT
' Prompts for a folder, loops every *Results*.xlsx in it,
' processes each into a new *_Formatted.xlsx file.
' =============================================================
Public Sub Main()

    ' ---- Folder picker ----
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select folder containing Journal Entry Results files"
    If fd.Show <> -1 Then Exit Sub

    Dim folderPath As String
    folderPath = fd.SelectedItems(1)

    ' ---- Collect matching filenames FIRST (Dir is NOT reentrant) ----
    ' We must build the full list before calling Workbooks.Open,
    ' because any nested Dir() call would reset the iterator.
    Dim files As New Collection
    Dim f As String
    f = Dir(folderPath & "\*Results*.xlsx")
    Do While f <> ""
        If InStr(LCase(f), "_formatted") = 0 Then   ' skip already-formatted files
            files.Add f
        End If
        f = Dir()
    Loop

    If files.Count = 0 Then
        MsgBox "No matching files found." & vbCrLf & _
               "Looking for .xlsx files with 'Results' in the name.", _
               vbInformation, "No Files Found"
        Exit Sub
    End If

    ' ---- Process each file ----
    Dim processedCount As Long
    Dim item As Variant
    For Each item In files
        ProcessFile folderPath, CStr(item)
        processedCount = processedCount + 1
    Next item

    MsgBox "Processing complete. " & processedCount & " file(s) formatted.", _
           vbInformation, "Done"
End Sub

' =============================================================
' PER-FILE ORCHESTRATOR
' Copies source → _Formatted.xlsx, applies all 11 transforms,
' saves as xlsx (strips macros), closes.
' =============================================================
Private Sub ProcessFile(folderPath As String, fileName As String)

    ' Declare wb at Sub level so the error handler can close it
    Dim wb As Workbook

    On Error GoTo ErrHandler

    ' Build paths
    Dim srcPath  As String : srcPath  = folderPath & "\" & fileName
    Dim baseName As String : baseName = Left(fileName, Len(fileName) - 5) ' strip ".xlsx"
    Dim dstPath  As String : dstPath  = folderPath & "\" & baseName & "_Formatted.xlsx"

    ' Remove stale output file (idempotency — safe to re-run)
    If Dir(dstPath) <> "" Then Kill dstPath
    FileCopy srcPath, dstPath

    ' Open the new copy; NEVER touch the original
    Set wb = Workbooks.Open(dstPath)

    ' Improve performance for large workbooks
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual
    Application.EnableEvents  = False

    ' ---- Transformations (order is significant — see plan) ----
    T3_CreatePBC                wb   ' must exist before T1 reorders
    T4_CreatePostClosing        wb   ' must exist before T1 reorders
    T5_CreateLegalFees          wb   ' reads TB24/TB25/GL; must exist before T1
    T6_ClassifyTestOpening      wb   ' inserts col C — must run before T8/T11
    T7_ClassifyTestCompleteness wb   ' inserts col C — must run before T8/T11
    T8_ReplaceNotFound          wb   ' depends on post-insert column positions
    T9_AppendSummaryRef         wb
    T10_HighlightTBComparison   wb
    T11_AppendConclusionAll     wb   ' sole appender of all conclusion blocks
    T2_ApplyTabColors           wb   ' cosmetic — order independent
    T1_ApplySheetOrder          wb   ' must be LAST (all sheets must already exist)

    ' Save as .xlsx (xlOpenXMLWorkbook = 51), which strips VBA from the output
    Application.DisplayAlerts = False
    wb.SaveAs dstPath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Set wb = Nothing

    ' Restore application state
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Application.EnableEvents  = True

    Exit Sub

ErrHandler:
    Dim errDesc As String : errDesc = Err.Description
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Application.EnableEvents  = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    MsgBox "Error processing '" & fileName & "':" & vbCrLf & errDesc, _
           vbExclamation, "Processing Error"
End Sub

' =============================================================
' T1 — SHEET ORDER
' Final order:
'   Summary, Test_Opening, Test_Completeness, Test_HighRisk,
'   Post closing JET Testing, TB_Comparison, Legal Fees Testing,
'   PBC >>>, TB24, TB25, GL
'
' Strategy: reverse-iterate the target array, moving each sheet
' to position 1 each time.  After the loop the array element at
' index 0 ends up at position 1 — exactly right, with no index
' drift.
' =============================================================
Private Sub T1_ApplySheetOrder(wb As Workbook)

    Dim order(10) As String
    order(0)  = "Summary"
    order(1)  = "Test_Opening"
    order(2)  = "Test_Completeness"
    order(3)  = "Test_HighRisk"
    order(4)  = "Post closing JET Testing"
    order(5)  = "TB_Comparison"
    order(6)  = "Legal Fees Testing"
    order(7)  = "PBC >>>"
    order(8)  = "TB24"
    order(9)  = "TB25"
    order(10) = "GL"

    Dim i As Integer
    For i = 10 To 0 Step -1
        On Error Resume Next
        wb.Sheets(order(i)).Move Before:=wb.Sheets(1)
        On Error GoTo 0
    Next i
End Sub

' =============================================================
' T2 — TAB COLORS
'   TB24, TB25, GL  → amber  RGB(255, 192,   0)
'   PBC >>>         → purple RGB(112,  48, 160)
'   All others      → no tab color
' =============================================================
Private Sub T2_ApplyTabColors(wb As Workbook)

    Dim nm As Variant

    ' Amber for source-data sheets
    For Each nm In Array("TB24", "TB25", "GL")
        On Error Resume Next
        wb.Sheets(CStr(nm)).Tab.Color = RGB(255, 192, 0)
        On Error GoTo 0
    Next nm

    ' Purple for the visual-divider sheet
    On Error Resume Next
    wb.Sheets("PBC >>>").Tab.Color = RGB(112, 48, 160)
    On Error GoTo 0

    ' No colour for all output / analysis sheets
    For Each nm In Array("Summary", "Test_Opening", "Test_Completeness", _
                         "Test_HighRisk", "Post closing JET Testing", _
                         "TB_Comparison", "Legal Fees Testing")
        On Error Resume Next
        wb.Sheets(CStr(nm)).Tab.ColorIndex = xlColorIndexNone
        On Error GoTo 0
    Next nm
End Sub

' =============================================================
' T3 — CREATE PBC >>> (empty visual divider sheet)
' Purple tab colour applied later in T2.
' =============================================================
Private Sub T3_CreatePBC(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "PBC >>>")
    ' Intentionally empty — serves only as a visual separator
End Sub

' =============================================================
' T4 — CREATE Post closing JET Testing
' Writes static content rows 1-23 only.
' The Conclusion & Findings block is appended by T11_PostClosing.
' =============================================================
Private Sub T4_CreatePostClosing(wb As Workbook)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "Post closing JET Testing")

    ws.Cells(1, 1).Value  = "kpmg"
    ws.Cells(2, 1).Value  = "Journal Entries Testing - Post-Closing Jes Test"
    ws.Cells(3, 1).Value  = "The purpose of reviewing post-closing journal entries in a financial audit is to detect errors, irregularities, and potential manipulation."
    ' Row 4: intentionally blank
    ws.Cells(5, 1).Value  = "Client Name"
    ws.Cells(5, 2).Value  = CLIENT_NAME
    ws.Cells(6, 1).Value  = "WP Ref"
    ws.Cells(6, 2).Value  = WP_REF
    ws.Cells(7, 1).Value  = "Period-end"
    ws.Cells(7, 2).Value  = PERIOD_END
    ' Rows 8-9: blank
    ws.Cells(10, 1).Value = "Post-Closing Journal Entries Criteria:"
    ws.Cells(11, 1).Value = "- We examine all Journal Entries which were posted after we have received the first version of the GL and TB as at year-end."
    ' Rows 12-13: blank
    ws.Cells(14, 1).Value = "Post-Closing JEs Test"
    ws.Cells(15, 1).Value = "- Not applicable as we have not received a new GL and/or TB for this Company. No further procedures performed."
    ' Rows 16-17: blank
    ws.Cells(18, 1).Value = "Findings"
    ws.Cells(19, 1).Value = "No findings."
    ' Rows 20-21: blank
    ws.Cells(22, 1).Value = "Conclusion"
    ws.Cells(23, 1).Value = "Satisfactory."
    ' Rows 24-25 remain blank.
    ' T11_PostClosing appends the standard Conclusion & Findings block at row 26.
End Sub

' =============================================================
' T5 — CREATE Legal Fees Testing
' Builds the sheet dynamically from TB24, TB25, and GL source data.
' The Conclusion & Findings block is appended by T11_LegalFees.
'
' Layout:
'   Rows  1- 8  : Header block (cols B/C/G)
'   Row  11     : "Lead Sheet" label
'   Row  12     : Column headers
'   Rows 13-14  : Accounts 85081 and 85082 (with live formulas)
'   Row  15     : Note row (<as per TB 2025> / <as per TB 2024>)
'   Rows 16-18  : Blank  (3-row gap)
'   Row  19     : GL transactions label
'   Row  20     : GL sub-headers
'   Rows 21+    : GL data rows (Konto = 85081 or 85082)
'   After data  : Total row, then Check row (+2)
' =============================================================
Private Sub T5_CreateLegalFees(wb As Workbook)

    Dim ws     As Worksheet : Set ws     = GetOrCreateSheet(wb, "Legal Fees Testing")
    Dim wsTB25 As Worksheet : Set wsTB25 = GetSheet(wb, "TB25")
    Dim wsTB24 As Worksheet : Set wsTB24 = GetSheet(wb, "TB24")
    Dim wsGL   As Worksheet : Set wsGL   = GetSheet(wb, "GL")

    ' ---- Header block (rows 1-8) ----
    ws.Cells(1, 2).Value = "ABCD"
    ws.Cells(1, 7).Value = "Legal fees"
    ws.Cells(2, 2).Value = "Client"
    ws.Cells(2, 7).Value = "Year end"
    ws.Cells(3, 2).Value = CLIENT_NAME
    ws.Cells(3, 7).Value = PERIOD_END
    ' Row 4: blank
    ws.Cells(5, 2).Value = "Objective:"
    ws.Cells(5, 3).Value = "To obtain evidence about the completeness of legal fees and the stated lawyers"
    ' Row 6: blank
    ws.Cells(7, 2).Value = "Procedures:"
    ws.Cells(7, 3).Value = "1) Obtain the breakdown of legal fees as per GL"
    ws.Cells(8, 3).Value = "2) To check if lawyers are in line with lawyer confirmation and request supporting documentation for other lawyers"
    ' Rows 9-10: blank

    ' ---- Lead Sheet table (rows 11-15) ----
    ws.Cells(11, 2).Value = "Lead Sheet"

    ws.Cells(12, 2).Value = "Account"
    ws.Cells(12, 3).Value = "Description"
    ws.Cells(12, 4).Value = "2025"
    ws.Cells(12, 5).Value = "2024"
    ws.Cells(12, 6).Value = "Difference"
    ws.Cells(12, 7).Value = "Difference in %"

    ' Row 13: account 85081
    Dim saldo85081_25 As Double : saldo85081_25 = LookupTBSaldo(wsTB25, "85081")
    Dim saldo85081_24 As Double : saldo85081_24 = LookupTBSaldo(wsTB24, "85081")
    Dim desc85081     As String
    desc85081 = LookupTBDescription(wsTB25, "85081")
    If desc85081 = "" Then desc85081 = LookupTBDescription(wsTB24, "85081")

    ws.Cells(13, 2).Value   = "85081"
    ws.Cells(13, 3).Value   = desc85081
    ws.Cells(13, 4).Value   = saldo85081_25
    ws.Cells(13, 5).Value   = saldo85081_24
    ws.Cells(13, 6).Formula = "=D13-E13"
    ws.Cells(13, 7).Formula = "=IF(E13=0,""-"",F13/E13)"

    ' Row 14: account 85082
    Dim saldo85082_25 As Double : saldo85082_25 = LookupTBSaldo(wsTB25, "85082")
    Dim saldo85082_24 As Double : saldo85082_24 = LookupTBSaldo(wsTB24, "85082")
    Dim desc85082     As String
    desc85082 = LookupTBDescription(wsTB25, "85082")
    If desc85082 = "" Then desc85082 = LookupTBDescription(wsTB24, "85082")

    ws.Cells(14, 2).Value   = "85082"
    ws.Cells(14, 3).Value   = desc85082
    ws.Cells(14, 4).Value   = saldo85082_25
    ws.Cells(14, 5).Value   = saldo85082_24
    ws.Cells(14, 6).Formula = "=D14-E14"
    ws.Cells(14, 7).Formula = "=IF(E14=0,""-"",F14/E14)"

    ' Row 15: note row
    ws.Cells(15, 4).Value = "<as per TB 2025>"
    ws.Cells(15, 5).Value = "<as per TB 2024>"

    ' ---- GL transactions block ----
    ' Rows 16-18 are blank (3-row gap).  Label row starts at row 19.
    Dim glLabelRow  As Long : glLabelRow  = 19
    Dim glDataStart As Long : glDataStart = glLabelRow + 2  ' row 21
    Dim glWriteRow  As Long : glWriteRow  = glDataStart

    ws.Cells(glLabelRow,     2).Value = "Transactions as per GL"
    ws.Cells(glLabelRow + 1, 2).Value = "Account"
    ws.Cells(glLabelRow + 1, 3).Value = "Date"
    ws.Cells(glLabelRow + 1, 4).Value = "Description"
    ws.Cells(glLabelRow + 1, 5).Value = "Amount"
    ws.Cells(glLabelRow + 1, 6).Value = "Comment"

    ' Read GL sheet into a Variant array for fast batch iteration
    If Not wsGL Is Nothing Then
        Dim glLastRow As Long : glLastRow = LastRow(wsGL, 6)  ' col F = Konto
        If glLastRow >= 2 Then
            Dim glArr As Variant
            glArr = wsGL.Range(wsGL.Cells(2, 1), wsGL.Cells(glLastRow, 9)).Value

            Dim r As Long
            For r = 1 To UBound(glArr, 1)
                Dim kontoVal As String
                kontoVal = Trim(CStr(glArr(r, 6)))  ' col F
                If kontoVal = "85081" Or kontoVal = "85082" Then
                    ws.Cells(glWriteRow, 2).Value = kontoVal
                    ws.Cells(glWriteRow, 3).Value = glArr(r, 3)                    ' Belegdatum  (col C)
                    ws.Cells(glWriteRow, 4).Value = glArr(r, 9)                    ' Buchungstext (col I)
                    ws.Cells(glWriteRow, 5).Value = ParseGermanNumber(glArr(r, 8)) ' Betrag       (col H)
                    ws.Cells(glWriteRow, 6).Value = "Ok, not material. No unexpected lawyer identified."
                    glWriteRow = glWriteRow + 1
                End If
            Next r
        End If
    End If

    ' Total row immediately after last data row
    Dim totalRow As Long : totalRow = glWriteRow
    ws.Cells(totalRow, 4).Value = "Total:"
    If glWriteRow > glDataStart Then
        ws.Cells(totalRow, 5).Formula = _
            "=SUM(E" & glDataStart & ":E" & (glWriteRow - 1) & ")"
    Else
        ws.Cells(totalRow, 5).Value = 0
    End If

    ' Check row: D14 (TB25 saldo 85082) minus the Total
    ' Placed 2 rows below Total (one blank row between)
    Dim checkRow As Long : checkRow = totalRow + 2
    ws.Cells(checkRow, 4).Value   = "Check"
    ws.Cells(checkRow, 5).Formula = "=D14-E" & totalRow

    ' T11_LegalFees appends the Conclusion & Findings block.
End Sub

' =============================================================
' T6 — Test_Opening: insert Classification column at col C
' =============================================================
Private Sub T6_ClassifyTestOpening(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Test_Opening")
    If ws Is Nothing Then Exit Sub
    AddClassificationColumn ws
End Sub

' =============================================================
' T7 — Test_Completeness: insert Classification column at col C
' =============================================================
Private Sub T7_ClassifyTestCompleteness(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Test_Completeness")
    If ws Is Nothing Then Exit Sub
    AddClassificationColumn ws
End Sub

' ---- Shared helper for T6 and T7 ----
' Inserts a new "Classification" column at col C (shifts existing C+ right).
' Applies:
'   • Solid blue (#4472C4) fill to the entire header row (row 1)
'   • Per data row: "PL" if Kontonummer starts with 6/7/8, else "BS"
'     BS rows → ThemeColor Accent1, tint +0.9  (light blue)
'     PL rows → ThemeColor Dark1,   tint -0.15 (light grey)
'   Fill is applied across col A to last used column in each row.
'
' Post-insert column layout for reference by T8 and T11:
'   Test_Opening:
'     A=Kontonummer B=Kontobezeichnung C=Classification D=TB24 Saldo
'     E=Opening Balance F=Difference G=Status H=Source
'   Test_Completeness:
'     A=Kontonummer B=Kontobezeichnung C=Classification D=Opening Balance
'     E=GL Movements F=Calculated Saldo G=TB25 Saldo H=Difference
'     I=Status J=Opening Source
' =============================================================
Private Sub AddClassificationColumn(ws As Worksheet)

    ' Idempotency: skip if Classification header is already in col C
    If Trim(CStr(ws.Cells(1, 3).Value)) = "Classification" Then Exit Sub

    ' Insert blank column at C; existing col C and beyond shift one right
    ws.Columns("C:C").Insert Shift:=xlToRight
    ws.Cells(1, 3).Value = "Classification"

    ' Apply solid blue fill to the entire header row (A1 : last-col-in-row-1)
    Dim hdrLastCol As Long
    hdrLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, hdrLastCol)).Interior
        .Pattern = xlSolid
        .Color   = RGB(68, 114, 196)   ' #4472C4
    End With

    ' Determine last data column (same as header — consistent across all rows)
    Dim dataLastCol As Long : dataLastCol = hdrLastCol

    Dim lastR As Long : lastR = LastRow(ws, 1)   ' anchor on col A (Kontonummer)

    Dim i As Long
    For i = 2 To lastR

        Dim acct As String : acct = Trim(CStr(ws.Cells(i, 1).Value))
        If acct = "" Then GoTo NextRow   ' skip blank / subtotal rows

        ' Classify
        Dim cls       As String
        Dim firstChar As String : firstChar = Left(acct, 1)
        If firstChar = "6" Or firstChar = "7" Or firstChar = "8" Then
            cls = "PL"
        Else
            cls = "BS"
        End If
        ws.Cells(i, 3).Value = cls

        ' Apply fill across the full row
        With ws.Range(ws.Cells(i, 1), ws.Cells(i, dataLastCol)).Interior
            If cls = "BS" Then
                .ThemeColor   = xlThemeColorAccent1
                .TintAndShade = 0.9
            Else
                .ThemeColor   = xlThemeColorDark1
                .TintAndShade = -0.15
            End If
        End With

NextRow:
    Next i
End Sub

' =============================================================
' T8 — Replace "NOT FOUND" with "TB24" in Source columns
'   Test_Opening    col H (8)  = Source        (after T6 insert)
'   Test_Completeness col J (10) = Opening Source (after T7 insert)
' Uses LookAt:=xlWhole for exact-match replacement (no substring hits).
' =============================================================
Private Sub T8_ReplaceNotFound(wb As Workbook)

    Dim wsO As Worksheet : Set wsO = GetSheet(wb, "Test_Opening")
    If Not wsO Is Nothing Then
        wsO.Columns(8).Replace What:="NOT FOUND", Replacement:="TB24", _
            LookAt:=xlWhole, MatchCase:=True
    End If

    Dim wsC As Worksheet : Set wsC = GetSheet(wb, "Test_Completeness")
    If Not wsC Is Nothing Then
        wsC.Columns(10).Replace What:="NOT FOUND", Replacement:="TB24", _
            LookAt:=xlWhole, MatchCase:=True
    End If
End Sub

' =============================================================
' T9 — Summary: append WP-reference sentence
' Leaves one blank row after the last existing content, then writes.
' Uses UsedRange to find the true last row (handles sparse content).
' =============================================================
Private Sub T9_AppendSummaryRef(wb As Workbook)

    Dim ws As Worksheet : Set ws = GetSheet(wb, "Summary")
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    ws.Cells(lastR + 2, 1).Value = _
        "Please refer to our procedure description of the Macro Process we used for the JE Testing. " & _
        "Please refer to <2.4.2.0020>."
End Sub

' =============================================================
' T10 — TB_Comparison: yellow highlight for "Only TB24" / "Only TB25"
' Applies solid fill #FFEB9C = RGB(255,235,156) across cols A-E.
' =============================================================
Private Sub T10_HighlightTBComparison(wb As Workbook)

    Dim ws As Worksheet : Set ws = GetSheet(wb, "TB_Comparison")
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long : lastR = LastRow(ws, 5)   ' col E = Status
    Dim i     As Long

    For i = 2 To lastR
        Dim statusVal As String
        statusVal = Trim(CStr(ws.Cells(i, 5).Value))
        If statusVal = "Only TB24" Or statusVal = "Only TB25" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 5)).Interior.Color = RGB(255, 235, 156)
        End If
    Next i
End Sub

' =============================================================
' T11 — CONCLUSION & FINDINGS (dispatcher)
' This is the SOLE place where Conclusion & Findings blocks are
' written.  T4 and T5 deliberately omit the block so T11 is
' authoritative across all 6 sheets.
' =============================================================
Private Sub T11_AppendConclusionAll(wb As Workbook)
    T11_TestOpening      wb
    T11_TestCompleteness wb
    T11_TestHighRisk     wb
    T11_PostClosing      wb
    T11_TBComparison     wb
    T11_LegalFees        wb
End Sub

' ---- Test_Opening ----
' TOTAL row at lastDataRow+1 (SUM cols D, E, F).
' Standard block at totalRow+3 (2 blank rows between).
Private Sub T11_TestOpening(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Test_Opening")
    If ws Is Nothing Then Exit Sub

    Dim lastR    As Long : lastR    = LastRow(ws, 1)  ' col A = Kontonummer
    Dim totalRow As Long : totalRow = lastR + 1

    ws.Cells(totalRow, 1).Value   = "TOTAL:"
    ' After T6: D=TB24 Saldo  E=Opening Balance  F=Difference
    ws.Cells(totalRow, 4).Formula = "=SUM(D2:D" & lastR & ")"
    ws.Cells(totalRow, 5).Formula = "=SUM(E2:E" & lastR & ")"
    ws.Cells(totalRow, 6).Formula = "=SUM(F2:F" & lastR & ")"

    AppendStandardBlock ws, totalRow + 3, 1
End Sub

' ---- Test_Completeness ----
' TOTAL row at lastDataRow+1 (SUM cols D-H).
' Standard block at totalRow+3.
Private Sub T11_TestCompleteness(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Test_Completeness")
    If ws Is Nothing Then Exit Sub

    Dim lastR    As Long : lastR    = LastRow(ws, 1)
    Dim totalRow As Long : totalRow = lastR + 1

    ws.Cells(totalRow, 1).Value   = "TOTAL:"
    ' After T7: D=Opening Balance  E=GL Movements  F=Calculated Saldo
    '           G=TB25 Saldo       H=Difference
    ws.Cells(totalRow, 4).Formula = "=SUM(D2:D" & lastR & ")"
    ws.Cells(totalRow, 5).Formula = "=SUM(E2:E" & lastR & ")"
    ws.Cells(totalRow, 6).Formula = "=SUM(F2:F" & lastR & ")"
    ws.Cells(totalRow, 7).Formula = "=SUM(G2:G" & lastR & ")"
    ws.Cells(totalRow, 8).Formula = "=SUM(H2:H" & lastR & ")"

    AppendStandardBlock ws, totalRow + 3, 1
End Sub

' ---- Test_HighRisk ----
' SUMMARY block 1 blank row after last data; standard block 2 rows after SUMMARY.
' Counts: total data rows, and rows where Gegenkonto (col D) is not 20000/49000.
Private Sub T11_TestHighRisk(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Test_HighRisk")
    If ws Is Nothing Then Exit Sub

    Dim lastR         As Long : lastR         = LastRow(ws, 1)
    Dim totalDataRows As Long : totalDataRows = lastR - 1  ' exclude header

    ' Count flagged entries
    Dim flagCount As Long : flagCount = 0
    Dim i         As Long
    For i = 2 To lastR
        Dim gk As String : gk = Trim(CStr(ws.Cells(i, 4).Value))
        If gk <> "20000" And gk <> "49000" Then flagCount = flagCount + 1
    Next i

    ' SUMMARY: 1 blank row after last data (lastR+2)
    Dim sumStart As Long : sumStart = lastR + 2
    ws.Cells(sumStart,     1).Value = "SUMMARY"
    ws.Cells(sumStart + 1, 1).Value = "Total entries for Konto 60000/60120: " & totalDataRows
    ws.Cells(sumStart + 2, 1).Value = "Flagged entries (Gegenkonto <> 20000 and <> 49000): " & flagCount

    ' Standard block: 2 blank rows after SUMMARY block → starts at sumStart+5
    AppendStandardBlock ws, sumStart + 5, 1
End Sub

' ---- Post closing JET Testing ----
' 2 blank rows after last content (row 23), then block at row 26.
Private Sub T11_PostClosing(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Post closing JET Testing")
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long : lastR = LastRow(ws, 1)  ' last content in col A = row 23
    AppendStandardBlock ws, lastR + 3, 1
End Sub

' ---- TB_Comparison ----
' 2 blank rows after last data row, standard block in col A.
Private Sub T11_TBComparison(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "TB_Comparison")
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long : lastR = LastRow(ws, 1)  ' col A = Kontonummer
    AppendStandardBlock ws, lastR + 3, 1
End Sub

' ---- Legal Fees Testing ----
' 2 blank rows after last content (the Check row), block in col B,
' with one extra lawyer-finding line after "No findings."
Private Sub T11_LegalFees(wb As Workbook)
    Dim ws As Worksheet : Set ws = GetSheet(wb, "Legal Fees Testing")
    If ws Is Nothing Then Exit Sub

    ' UsedRange is reliable here — Legal Fees is a freshly created sheet
    Dim lastR As Long
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    ' colOffset=2 (col B); extra line after "No findings."
    AppendStandardBlock ws, lastR + 3, 2, _
        "We did not identify any lawyer that has not been represented to us by Management."
End Sub

' =============================================================
' HELPER FUNCTIONS
' =============================================================

' GetOrCreateSheet
' Returns a fresh worksheet named sheetName appended at the end.
' If a sheet with that name already exists it is deleted first
' (ensures idempotency when re-running on the same output file).
Private Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = sheetName
    Set GetOrCreateSheet = ws
End Function

' GetSheet
' Returns the named sheet, or Nothing if it does not exist (no error raised).
Private Function GetSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

' LastRow
' Returns the last row that has data in the specified column.
Private Function LastRow(ws As Worksheet, col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ParseGermanNumber
' Converts a cell value that may be stored as a German-formatted string
' (period = thousands separator, comma = decimal separator) to a Double.
' If the value is already a numeric VBA type it is returned directly,
' avoiding false stripping of a decimal point in the English representation.
Private Function ParseGermanNumber(rawVal As Variant) As Double
    If IsEmpty(rawVal) Or CStr(rawVal) = "" Then
        ParseGermanNumber = 0
        Exit Function
    End If

    Select Case VarType(rawVal)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbByte, vbDecimal, vbCurrency
            ParseGermanNumber = CDbl(rawVal)
            Exit Function
    End Select

    Dim s As String : s = CStr(rawVal)
    s = Replace(s, ".", "")    ' strip thousands separator (period)
    s = Replace(s, ",", ".")   ' convert decimal comma to dot

    If IsNumeric(s) Then
        ParseGermanNumber = CDbl(s)
    Else
        ParseGermanNumber = 0
    End If
End Function

' LookupTBSaldo
' Searches col A of a TB sheet (data from row 3) for kontonummer and
' returns the parsed value from col H (Saldo).  Returns 0 if not found.
Private Function LookupTBSaldo(wsTB As Worksheet, kontonummer As String) As Double
    If wsTB Is Nothing Then LookupTBSaldo = 0 : Exit Function
    Dim lastR As Long : lastR = LastRow(wsTB, 1)
    Dim i As Long
    For i = 3 To lastR
        If Trim(CStr(wsTB.Cells(i, 1).Value)) = kontonummer Then
            LookupTBSaldo = ParseGermanNumber(wsTB.Cells(i, 8).Value)
            Exit Function
        End If
    Next i
    LookupTBSaldo = 0
End Function

' LookupTBDescription
' Searches col A of a TB sheet for kontonummer and returns col B (Kontobezeichnung).
Private Function LookupTBDescription(wsTB As Worksheet, kontonummer As String) As String
    If wsTB Is Nothing Then LookupTBDescription = "" : Exit Function
    Dim lastR As Long : lastR = LastRow(wsTB, 1)
    Dim i As Long
    For i = 3 To lastR
        If Trim(CStr(wsTB.Cells(i, 1).Value)) = kontonummer Then
            LookupTBDescription = CStr(wsTB.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    LookupTBDescription = ""
End Function

' AppendStandardBlock
' Writes the standard Conclusion & Findings block starting at startRow.
'   colOffset : column index (1=A, 2=B, …)
'   extraLine : optional additional line written after "No findings."
'               (used by Legal Fees Testing for the lawyer sentence)
'
' Block layout:
'   startRow+0  Conclusion:
'   startRow+1  No issues noted.
'   startRow+2  (blank)
'   startRow+3  Findings
'   startRow+4  No findings.
'  [startRow+5  extraLine, if provided]
Private Sub AppendStandardBlock(ws As Worksheet, startRow As Long, _
                                 colOffset As Long, _
                                 Optional extraLine As String = "")
    ws.Cells(startRow,     colOffset).Value = "Conclusion:"
    ws.Cells(startRow + 1, colOffset).Value = "No issues noted."
    ' startRow+2: blank
    ws.Cells(startRow + 3, colOffset).Value = "Findings"
    ws.Cells(startRow + 4, colOffset).Value = "No findings."
    If extraLine <> "" Then
        ws.Cells(startRow + 5, colOffset).Value = extraLine
    End If
End Sub
