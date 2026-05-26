Attribute VB_Name = "BSPLBalanceSums"
Option Explicit

' =============================================================
' CONSTANTS
' Edit these to match the workbook layout.
' =============================================================
Private Const ROW_ENTITY      As Long = 9    ' row containing entity codes (e.g. "1234")
Private Const ROW_DATA_LABEL  As Long = 10   ' row where "DATA" marks the comparison column
Private Const ROW_DATA_START  As Long = 13   ' first row that may contain account numbers
Private Const COL_ACCOUNT     As Long = 2    ' column B holds account numbers in BS / PL

Private Const TB_COL_ACCT     As Long = 1    ' TB sheet: account code in col A
Private Const TB_COL_SALDO    As Long = 6    ' TB sheet: ending balance in col F
Private Const TB_ROW_START    As Long = 2    ' TB sheet: data starts at row 2 (row 1 = header)

' Entity code to exclude entirely (sheet will not be processed).
Private Const EXCLUDED_CODE   As String = "1224"

' =============================================================
' MODULE-LEVEL STATE
' =============================================================
Private m_SkipLog As String

' =============================================================
' UDT — one entity column found in ROW_ENTITY
' =============================================================
Private Type EntityInfo
    Code        As String   ' 4-digit entity code (also the TB sheet name)
    FirstCol    As Long     ' leftmost column of entity span in ROW_ENTITY
    LastCol     As Long     ' rightmost column of entity span
    DataCol     As Long     ' column in ROW_DATA_LABEL that contains "DATA" (value to compare)
    InsertCol   As Long     ' column where "as per TB" will be inserted
    Skipped     As Boolean
End Type

' =============================================================
' PUBLIC ENTRY POINT
' Runs the balance-sum / difference / row-hide logic on both
' the "BS" and "PL" sheets in the active workbook.
' The TB sheets must already be present and named by their
' 4-digit entity code (exactly as it appears in ROW_ENTITY).
' =============================================================
Public Sub RunBSPLBalanceSums()

    Dim sheetNames As Variant
    Dim item       As Variant
    Dim ws         As Worksheet
    Dim processed  As Long

    m_SkipLog = ""

    sheetNames = Array("BS", "PL")

    AppPerfOn2

    On Error GoTo ErrHandler

    For Each item In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(item))
        On Error GoTo ErrHandler
        If ws Is Nothing Then
            LogSkip2 "Sheet '" & CStr(item) & "' not found in this workbook — skipped."
        Else
            ProcessConsoSheet ws
            processed = processed + 1
        End If
    Next item

CleanExit:
    AppPerfOff2
    If m_SkipLog <> "" Then
        MsgBox "Done (" & processed & " sheet(s) processed) with warnings:" & _
               vbCrLf & vbCrLf & m_SkipLog, vbExclamation, "BS/PL Balance Sums — Warnings"
    Else
        MsgBox "Done. " & processed & " sheet(s) processed.", _
               vbInformation, "BS/PL Balance Sums"
    End If
    Exit Sub

ErrHandler:
    Dim errMsg As String : errMsg = Err.Description
    AppPerfOff2
    MsgBox "Unexpected error:" & vbCrLf & errMsg, vbExclamation, "BS/PL Balance Sums — Error"

End Sub

' =============================================================
' PER-SHEET PROCESSOR
' =============================================================
Private Sub ProcessConsoSheet(ws As Worksheet)

    Dim entities() As EntityInfo
    Dim entCount   As Long
    Dim i          As Long
    Dim wsTB       As Worksheet

    FindEntityColumns2 ws, entities, entCount

    If entCount = 0 Then
        LogSkip2 "Sheet '" & ws.Name & "': no entity codes found in row " & ROW_ENTITY & " — skipped."
        Exit Sub
    End If

    ' Process RIGHT TO LEFT to prevent column-index drift on insertion.
    For i = entCount - 1 To 0 Step -1
        If Not entities(i).Skipped Then
            Set wsTB = Nothing
            On Error Resume Next
            Set wsTB = ThisWorkbook.Sheets(entities(i).Code)
            On Error GoTo 0
            If wsTB Is Nothing Then
                LogSkip2 "Sheet '" & ws.Name & "', entity '" & entities(i).Code & _
                         "': TB sheet not found — skipped."
            Else
                InsertAndFill ws, entities(i), wsTB
            End If
        End If
    Next i

End Sub

' =============================================================
' CORE: INSERT COLUMNS AND FILL
' For each entity:
'   1. Insert "as per TB" + "Difference" columns (idempotent).
'   2. Walk account clusters (consecutive account rows separated
'      by blank rows; blue section-header rows are skipped).
'   3. First row of each cluster: write cluster TB sum and
'      difference vs. entity reported value.
'   4. Hide detail rows (rows 2..n of each cluster).
' =============================================================
Private Sub InsertAndFill(ws As Worksheet, ent As EntityInfo, wsTB As Worksheet)

    Dim insertAt     As Long
    Dim diffAt       As Long
    Dim entityCol    As Long
    Dim lastDataRow  As Long
    Dim r            As Long
    Dim r2           As Long
    Dim acctVal      As String
    Dim nextAcct     As String
    Dim clusterTotal As Double
    Dim entityVal    As Double

    insertAt  = ent.InsertCol
    entityCol = ent.DataCol   ' the "DATA" column — stays fixed, insertions are to its right

    ' ---- Idempotency: remove previously inserted columns ----
    Do While Trim(CStr(ws.Cells(ROW_ENTITY, insertAt).Value)) = "as per TB" _
          Or Trim(CStr(ws.Cells(ROW_ENTITY, insertAt).Value)) = "Difference"
        Application.DisplayAlerts = False
        ws.Columns(insertAt).Delete Shift:=xlToLeft
        Application.DisplayAlerts = True
    Loop

    lastDataRow = LastUsedRow2(ws, COL_ACCOUNT)

    ' ---- Unhide all data rows (reset from any previous run) ----
    If lastDataRow >= ROW_DATA_START Then
        ws.Rows(ROW_DATA_START & ":" & lastDataRow).Hidden = False
    End If

    ' ---- Insert "as per TB" then "Difference" ----
    ws.Columns(insertAt).Insert Shift:=xlToRight
    ws.Cells(ROW_ENTITY, insertAt).Value = "as per TB"

    diffAt = insertAt + 1
    ws.Columns(diffAt).Insert Shift:=xlToRight
    ws.Cells(ROW_ENTITY, diffAt).Value = "Difference"

    ' ---- Walk data rows: detect clusters, sum, fill, hide ----
    r = ROW_DATA_START
    Do While r <= lastDataRow

        ' Skip blue section-header rows entirely (don't treat them as cluster boundary).
        If IsSkipRow(ws, r) Then
            r = r + 1
        ElseIf IsAccountRow2(ws.Cells(r, COL_ACCOUNT).Value) Then
            ' --- First account of a new cluster ---
            ' CleanAccount strips the leading apostrophe Excel stores as text-prefix.
            acctVal = CleanAccount(ws.Cells(r, COL_ACCOUNT).Value)
            clusterTotal = SumAccountInTB2(wsTB, acctVal)
            r2 = r + 1

            ' Collect remaining accounts in this cluster.
            ' A blue section-header row OR any non-account row ends the cluster.
            ' Previously blue rows were skipped inside the loop, which caused the
            ' entire sheet to be treated as one giant cluster.
            Do While r2 <= lastDataRow
                If IsSkipRow(ws, r2) Then Exit Do
                If Not IsAccountRow2(ws.Cells(r2, COL_ACCOUNT).Value) Then Exit Do
                nextAcct = CleanAccount(ws.Cells(r2, COL_ACCOUNT).Value)
                clusterTotal = clusterTotal + SumAccountInTB2(wsTB, nextAcct)
                r2 = r2 + 1
            Loop

            ' Write TB total and difference in the first row of the cluster.
            ws.Cells(r, insertAt).Value = clusterTotal
            entityVal = ParseGermanNumber2(ws.Cells(r, entityCol).Value)
            ws.Cells(r, diffAt).Value = entityVal - clusterTotal

            ' Hide detail rows (every row in the cluster except the first).
            If r2 - 1 > r Then
                ws.Rows(r + 1 & ":" & (r2 - 1)).Hidden = True
            End If

            r = r2
        Else
            r = r + 1   ' blank or non-account row → cluster boundary, keep scanning
        End If

    Loop

End Sub

' =============================================================
' DISCOVERY
' =============================================================

' FindEntityColumns2
' Scans ROW_ENTITY from col 3 onwards (cols A/B are account data).
' Uses MergeArea.Columns.Count to handle merged entity headers.
' Skips the excluded code (EXCLUDED_CODE constant).
' For each entity, finds the column in ROW_DATA_LABEL that contains
' "DATA" — that column is stored as DataCol and used for comparisons.
' Entities with no "DATA" column in ROW_DATA_LABEL are marked Skipped.
Private Sub FindEntityColumns2(ws As Worksheet, _
                                ByRef entities() As EntityInfo, _
                                ByRef entCount As Long)

    Dim lastCol  As Long
    Dim c        As Long
    Dim dc       As Long
    Dim cellVal  As String
    Dim spanCols As Long
    Dim dataCol  As Long

    entCount = 0
    ReDim entities(0)

    lastCol = ws.Cells(ROW_ENTITY, ws.Columns.Count).End(xlToLeft).Column

    c = 3   ' cols A and B are account data — start scanning from col C
    Do While c <= lastCol
        cellVal = Trim(CStr(ws.Cells(ROW_ENTITY, c).Value))
        If cellVal <> "" Then
            spanCols = ws.Cells(ROW_ENTITY, c).MergeArea.Columns.Count

            ' Skip explicitly excluded entity codes.
            If cellVal = EXCLUDED_CODE Then
                c = c + spanCols
            Else
                ' Find the "DATA" column within this entity's span in ROW_DATA_LABEL.
                ' If no "DATA" label is found, skip this entity entirely.
                dataCol = 0
                For dc = c To c + spanCols - 1
                    If Trim(CStr(ws.Cells(ROW_DATA_LABEL, dc).Value)) = "DATA" Then
                        dataCol = dc
                        Exit For
                    End If
                Next dc

                ReDim Preserve entities(entCount)
                entities(entCount).Code      = cellVal
                entities(entCount).FirstCol  = c
                entities(entCount).LastCol   = c + spanCols - 1
                entities(entCount).DataCol   = dataCol
                entities(entCount).InsertCol = c + spanCols
                ' Mark as skipped when no DATA column found — no columns will be inserted.
                entities(entCount).Skipped   = (dataCol = 0)
                entCount = entCount + 1

                c = c + spanCols
            End If
        Else
            c = c + 1
        End If
    Loop

End Sub

' =============================================================
' DATA FUNCTIONS
' =============================================================

' SumAccountInTB2
' Sums col F (TB_COL_SALDO) of the TB sheet for all rows where
' col A (TB_COL_ACCT) exactly matches accountCode.
Private Function SumAccountInTB2(wsTB As Worksheet, accountCode As String) As Double

    Dim lastR As Long
    Dim arr   As Variant
    Dim total As Double
    Dim r     As Long

    If wsTB Is Nothing Then SumAccountInTB2 = 0 : Exit Function

    lastR = LastUsedRow2(wsTB, TB_COL_ACCT)
    If lastR < TB_ROW_START Then SumAccountInTB2 = 0 : Exit Function

    arr = wsTB.Range(wsTB.Cells(TB_ROW_START, 1), _
                     wsTB.Cells(lastR, TB_COL_SALDO)).Value

    total = 0
    For r = 1 To UBound(arr, 1)
        If Trim(CStr(arr(r, TB_COL_ACCT))) = accountCode Then
            total = total + ParseGermanNumber2(arr(r, TB_COL_SALDO))
        End If
    Next r

    SumAccountInTB2 = total

End Function

' =============================================================
' UTILITY FUNCTIONS
' =============================================================

' CleanAccount
' Strips the leading apostrophe that Excel stores when a number is
' forced to text (e.g. '1140013400 → 1140013400).
Private Function CleanAccount(cellVal As Variant) As String
    Dim s As String
    s = Trim(CStr(cellVal))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    CleanAccount = Trim(s)
End Function

' IsAccountRow2
' Returns True if the cell value looks like an account number:
' starts with a digit (0-9) after stripping any leading apostrophe.
' Works for both hyphenated codes ("1111-2050") and plain numeric
' codes ("1140013400").
Private Function IsAccountRow2(cellVal As Variant) As Boolean
    Dim s As String
    s = CleanAccount(cellVal)
    If Len(s) = 0 Then IsAccountRow2 = False : Exit Function
    IsAccountRow2 = (s Like "[0-9]*")
End Function

' IsSkipRow
' Returns True for coloured section-header rows (the blue rows visible
' in the screenshot).  Any row whose account cell has a non-white,
' non-transparent background fill is treated as a header to skip.
' ColorIndex -4142 (xlColorIndexNone) = no fill; 2 = white — both
' appear white and should NOT be skipped.
Private Function IsSkipRow(ws As Worksheet, r As Long) As Boolean
    Dim ci As Long
    ci = ws.Cells(r, COL_ACCOUNT).Interior.ColorIndex
    IsSkipRow = (ci <> xlColorIndexNone) And (ci <> 2)
End Function

' LastUsedRow2
' Returns the last row with data in the given column.
Private Function LastUsedRow2(ws As Worksheet, col As Long) As Long
    LastUsedRow2 = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ParseGermanNumber2
' Converts German-formatted numeric strings (. = thousands, , = decimal)
' or already-numeric values to Double.
Private Function ParseGermanNumber2(rawVal As Variant) As Double
    Dim s As String

    If IsEmpty(rawVal) Or CStr(rawVal) = "" Then
        ParseGermanNumber2 = 0 : Exit Function
    End If

    Select Case VarType(rawVal)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbByte, vbDecimal, vbCurrency
            ParseGermanNumber2 = CDbl(rawVal)
            Exit Function
    End Select

    s = CStr(rawVal)
    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")

    If IsNumeric(s) Then
        ParseGermanNumber2 = CDbl(s)
    Else
        ParseGermanNumber2 = 0
    End If
End Function

' LogSkip2
' Accumulates warning messages shown in one MsgBox at the end.
Private Sub LogSkip2(msg As String)
    If m_SkipLog = "" Then
        m_SkipLog = msg
    Else
        m_SkipLog = m_SkipLog & vbCrLf & msg
    End If
End Sub

' AppPerfOn2 / AppPerfOff2
' Toggle Excel performance settings for the duration of the run.
Private Sub AppPerfOn2()
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual
    Application.EnableEvents  = False
End Sub

Private Sub AppPerfOff2()
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Application.EnableEvents  = True
    Application.DisplayAlerts = True
End Sub
