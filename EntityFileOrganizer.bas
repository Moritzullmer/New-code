Attribute VB_Name = "EntityFileOrganizer"
Option Explicit

' =============================================================
' ENTITY FILE ORGANIZER
'
' Reads a mapping workbook listing entities with TB codes and GL
' names, scans a source root whose immediate subfolders are
' entity folders (each with TB 24/ and TB 25/ subfolders), and
' filters one shared consolidated GL workbook per entity.  For
' every mapping row it produces an output folder named after the
' entity, with three subfolders:
'   <entity>\TB 24\  - the matched TB 24 file (copied unchanged)
'   <entity>\TB 25\  - the matched TB 25 file (copied unchanged)
'   <entity>\GL\GL.xlsx
'                    - the consolidated GL restricted to rows
'                      whose column H equals one of that row's
'                      mapping GL names (header row always kept).
'
' Mapping layout on sheet 1 of the mapping workbook:
'   Row 3 headers:   A=Entity Name  B=TB #   C=GL 1  D=GL 2  E=GL 3 ...
'   Row 4+ data:     one row per entity
'   GL columns:      read until the row-3 header cell is empty
'
' Matching rules:
'   - TB file:  cell A1 equals the row's TB # (Trim, case-insensitive).
'   - GL rows:  column H equals one of the row's GL names
'               (Trim, case-insensitive).
' =============================================================

' =============================================================
' PUBLIC ENTRY POINT
' Prompts for four paths (mapping workbook, source root,
' consolidated GL workbook, output root) and builds one output
' subfolder per mapping row, each with TB 24\, TB 25\, GL\
' subfolders holding the matched files.
' =============================================================
Public Sub Main()

    ' ---- Pickers ----
    Dim mappingPath As String
    mappingPath = PickFile("Select the mapping workbook")
    If mappingPath = "" Then Exit Sub

    Dim sourceRoot As String
    sourceRoot = PickFolder("Select the source root folder (contains entity subfolders)")
    If sourceRoot = "" Then Exit Sub

    Dim glPath As String
    glPath = PickFile("Select the consolidated GL workbook")
    If glPath = "" Then Exit Sub

    Dim outputRoot As String
    outputRoot = PickFolder("Select the output root folder")
    If outputRoot = "" Then Exit Sub

    ' ---- Speed up ----
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim mappingRows As Collection : Set mappingRows = New Collection
    Dim warnings As Collection : Set warnings = New Collection
    Dim okCount As Long : okCount = 0
    Dim errDesc As String : errDesc = ""

    On Error GoTo Cleanup

    ' ---- Read the mapping ----
    Dim nGL As Long
    ReadMappingRows mappingPath, mappingRows, nGL
    If mappingRows.Count = 0 Then
        warnings.Add "Mapping is empty (no data rows from row 4 onwards in column A)."
        GoTo Cleanup
    End If

    ' ---- Pre-scan TB files (one sweep across both years) ----
    Dim tbIndex As Object
    Set tbIndex = BuildTBIndex(sourceRoot)

    ' ---- Load the consolidated GL once into memory ----
    Dim glData As Variant   ' 1-based 2D array; empty Variant if load failed
    Dim glLoadMsg As String
    LoadConsolidatedGL glPath, glData, glLoadMsg
    If glLoadMsg <> "" Then warnings.Add "Consolidated GL: " & glLoadMsg

    ' ---- Process each mapping row ----
    Dim i As Long
    For i = 1 To mappingRows.Count
        Dim mapRow As Object : Set mapRow = mappingRows(i)
        Dim entityName As String : entityName = CStr(mapRow("EntityName"))
        Dim tbCode As String : tbCode = CStr(mapRow("TBCode"))
        Dim glNames As Variant : glNames = mapRow("GLNames")

        Dim rowWarn As String : rowWarn = ""

        If tbCode = "" Then
            warnings.Add entityName & ": TB # is empty in mapping; skipped."
            GoTo NextRow
        End If
        If Not tbIndex.Exists(tbCode) Then
            warnings.Add entityName & ": TB code """ & tbCode & _
                         """ not found in any entity folder; skipped."
            GoTo NextRow
        End If

        Dim rec As Object : Set rec = tbIndex(tbCode)
        Dim entityFolder As String : entityFolder = CStr(rec("folder"))
        Dim tb24File As String : tb24File = CStr(rec("tb24"))
        Dim tb25File As String : tb25File = CStr(rec("tb25"))

        Dim sanitized As String : sanitized = SanitizeFolderName(entityName)
        If sanitized = "" Then
            warnings.Add entityName & ": entity name sanitized to empty; skipped."
            GoTo NextRow
        End If

        Dim outFolder As String : outFolder = outputRoot & "\" & sanitized
        If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder
        Dim outTB24 As String : outTB24 = outFolder & "\TB 24"
        Dim outTB25 As String : outTB25 = outFolder & "\TB 25"
        Dim outGL As String : outGL = outFolder & "\GL"
        If Dir(outTB24, vbDirectory) = "" Then MkDir outTB24
        If Dir(outTB25, vbDirectory) = "" Then MkDir outTB25
        If Dir(outGL, vbDirectory) = "" Then MkDir outGL

        ' --- TB 24 copy ---
        If tb24File = "" Then
            rowWarn = rowWarn & "TB 24 missing; "
        Else
            Dim tb24Dst As String : tb24Dst = outTB24 & "\" & tb24File
            On Error Resume Next
            If Dir(tb24Dst) <> "" Then Kill tb24Dst
            Err.Clear
            FileCopy entityFolder & "\TB 24\" & tb24File, tb24Dst
            If Err.Number <> 0 Then
                rowWarn = rowWarn & "TB 24 copy failed (" & Err.Description & "); "
                Err.Clear
            End If
            On Error GoTo Cleanup
        End If

        ' --- TB 25 copy ---
        If tb25File = "" Then
            rowWarn = rowWarn & "TB 25 missing; "
        Else
            Dim tb25Dst As String : tb25Dst = outTB25 & "\" & tb25File
            On Error Resume Next
            If Dir(tb25Dst) <> "" Then Kill tb25Dst
            Err.Clear
            FileCopy entityFolder & "\TB 25\" & tb25File, tb25Dst
            If Err.Number <> 0 Then
                rowWarn = rowWarn & "TB 25 copy failed (" & Err.Description & "); "
                Err.Clear
            End If
            On Error GoTo Cleanup
        End If

        ' --- GL filter (from the consolidated GL loaded in memory) ---
        Dim mergedPath As String : mergedPath = outGL & "\GL.xlsx"
        On Error Resume Next
        If Dir(mergedPath) <> "" Then Kill mergedPath
        On Error GoTo Cleanup

        Dim expectedGL As Long : expectedGL = CountNonEmpty(glNames)
        If expectedGL = 0 Then
            rowWarn = rowWarn & "no GL names in mapping; "
        ElseIf IsEmpty(glData) Then
            rowWarn = rowWarn & "consolidated GL not loaded; skipping filter; "
        Else
            Dim keptRows As Long : keptRows = 0
            On Error Resume Next
            keptRows = FilterConsolidatedGL(glData, glNames, mergedPath)
            If Err.Number <> 0 Then
                rowWarn = rowWarn & "GL filter failed (" & Err.Description & "); "
                Err.Clear
            End If
            On Error GoTo Cleanup
            If keptRows = 0 Then
                rowWarn = rowWarn & "no GL rows matched column H for this entity; "
            End If
        End If

        okCount = okCount + 1
        If rowWarn <> "" Then warnings.Add entityName & ": " & rowWarn
NextRow:
    Next i

Cleanup:
    If Err.Number <> 0 Then errDesc = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim totalRows As Long
    If mappingRows Is Nothing Then totalRows = 0 Else totalRows = mappingRows.Count

    Dim summary As String : summary = ""
    If errDesc <> "" Then summary = "Unexpected error: " & errDesc & vbCrLf & vbCrLf
    summary = summary & okCount & " of " & totalRows & " entities processed."
    If warnings.Count > 0 Then
        summary = summary & vbCrLf & vbCrLf & "Warnings (" & warnings.Count & "):" & vbCrLf
        Dim w As Variant
        For Each w In warnings
            summary = summary & "  - " & CStr(w) & vbCrLf
        Next w
    End If

    Dim icon As VbMsgBoxStyle
    If errDesc <> "" Or warnings.Count > 0 Then icon = vbExclamation Else icon = vbInformation
    MsgBox summary, icon, "Entity File Organizer"
End Sub

' =============================================================
' CountNonEmpty
' Counts the non-blank (after Trim) entries in a Variant array.
' =============================================================
Private Function CountNonEmpty(arr As Variant) As Long
    Dim n As Long : n = 0
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Trim(CStr(arr(i))) <> "" Then n = n + 1
    Next i
    CountNonEmpty = n
End Function

' =============================================================
' DIALOG HELPERS
' Return "" when the user cancels, so the caller can Exit Sub.
' =============================================================
Private Function PickFile(title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = title
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
    If fd.Show = -1 Then
        PickFile = fd.SelectedItems(1)
    Else
        PickFile = ""
    End If
End Function

Private Function PickFolder(title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = title
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If
End Function

' =============================================================
' SanitizeFolderName
' Replaces Windows-invalid characters ( \ / : * ? " < > | ) with
' underscores and strips trailing spaces/dots (also invalid).
' Leaves the rest of the entity name unchanged so e.g. apostrophes
' and parentheses survive ("St Martin's Tower Propco").
' =============================================================
Private Function SanitizeFolderName(nameRaw As String) As String
    Dim s As String : s = nameRaw
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(bad), "_")
    Next bad
    ' Windows disallows trailing spaces or dots in folder names
    Do While Len(s) > 0 And (Right(s, 1) = " " Or Right(s, 1) = ".")
        s = Left(s, Len(s) - 1)
    Loop
    SanitizeFolderName = Trim(s)
End Function

' =============================================================
' AddFilesByPattern
' Adds every file matching `folder\pattern` to `coll`.
' Self-contained: the Dir() iterator starts and ends inside the
' Do loop so callers can safely chain multiple patterns without
' tripping Dir's non-reentrancy.
' =============================================================
Private Sub AddFilesByPattern(coll As Collection, folder As String, pattern As String)
    Dim f As String : f = Dir(folder & "\" & pattern)
    Do While f <> ""
        coll.Add f
        f = Dir()
    Loop
End Sub

' =============================================================
' CollectSubfolders
' Returns a Collection of full paths for every immediate
' subdirectory of `parentPath` (skipping "." and "..").
' =============================================================
Private Function CollectSubfolders(parentPath As String) As Collection
    Dim result As New Collection
    Dim f As String : f = Dir(parentPath & "\*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            Dim fp As String : fp = parentPath & "\" & f
            If (GetAttr(fp) And vbDirectory) = vbDirectory Then
                result.Add fp
            End If
        End If
        f = Dir()
    Loop
    Set CollectSubfolders = result
End Function

' =============================================================
' LastUsedRow / LastUsedCol
' Use Find to locate the last non-empty cell in a worksheet.
' Returns 0 if the sheet is completely empty (Find returns
' Nothing and the On Error swallows the member-access error).
' =============================================================
Private Function LastUsedRow(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then LastUsedRow = 0 Else LastUsedRow = c.Row
End Function

Private Function LastUsedCol(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then LastUsedCol = 0 Else LastUsedCol = c.Column
End Function

' =============================================================
' ReadMappingRows
' Opens the mapping workbook read-only and reads the entity table
' from sheet 1.
'
' Row 3 is the header row: A=Entity Name  B=TB #  C=GL 1  D=GL 2 ...
' The GL column count is derived by walking row 3 from col C
' rightwards until an empty header cell is hit.
'
' Row 4+ is data: one entity per row, read until col A is empty.
' Each row is returned as a Scripting.Dictionary with keys:
'   "EntityName" (String)
'   "TBCode"     (String)
'   "GLNames"    (Variant array, 1..nGL; "" where a cell was blank)
'
' If the mapping workbook is already open in the current Excel
' instance the existing Workbook object is reused and NOT closed
' after reading, so running the macro from the mapping workbook
' itself is safe.
' =============================================================
Private Sub ReadMappingRows(path As String, ByRef rows As Collection, ByRef nGL As Long)
    Set rows = New Collection
    nGL = 0

    Dim nameOnly As String : nameOnly = Mid(path, InStrRev(path, "\") + 1)
    Dim wb As Workbook
    Dim alreadyOpen As Boolean

    On Error Resume Next
    Set wb = Workbooks(nameOnly)
    On Error GoTo 0
    alreadyOpen = Not (wb Is Nothing)
    If Not alreadyOpen Then
        Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    End If

    Dim ws As Worksheet : Set ws = wb.Sheets(1)

    ' Count GL columns starting at col C (3)
    Dim c As Long : c = 3
    Do While Trim(CStr(ws.Cells(3, c).Value)) <> ""
        nGL = nGL + 1
        c = c + 1
    Loop

    ' Read entity rows from row 4 while col A is non-empty
    Dim r As Long : r = 4
    Do While Trim(CStr(ws.Cells(r, 1).Value)) <> ""
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = 1 ' vbTextCompare
        rec("EntityName") = Trim(CStr(ws.Cells(r, 1).Value))
        rec("TBCode") = Trim(CStr(ws.Cells(r, 2).Value))

        Dim glArr As Variant
        If nGL > 0 Then
            ReDim glArr(1 To nGL)
            Dim k As Long
            For k = 1 To nGL
                glArr(k) = Trim(CStr(ws.Cells(r, 2 + k).Value))
            Next k
        Else
            ReDim glArr(1 To 1)
            glArr(1) = ""
        End If
        rec("GLNames") = glArr

        rows.Add rec
        r = r + 1
    Loop

    If Not alreadyOpen Then wb.Close SaveChanges:=False
End Sub

' =============================================================
' BuildTBIndex
' One pass across every immediate subfolder of sourceRoot:
'   - For each file in <entity>\TB 24\ : read A1, register tb24.
'   - For each file in <entity>\TB 25\ : read A1, register tb25.
'
' Returns a Scripting.Dictionary keyed case-insensitively by the
' trimmed A1 text.  The value is itself a Dictionary with keys:
'   "folder" : full path of the source entity folder
'   "tb24"   : TB 24 filename (or "" if not seen)
'   "tb25"   : TB 25 filename (or "" if not seen)
'
' A code first seen in one year creates the record; the other
' year's pass fills the remaining slot on the same record when it
' encounters the same A1 value.  If a year never yields a match,
' that slot stays "" so the row processor can emit a precise
' "TB 24 missing" / "TB 25 missing" warning.
' =============================================================
Private Function BuildTBIndex(sourceRoot As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare

    Dim entityFolders As Collection
    Set entityFolders = CollectSubfolders(sourceRoot)

    Dim ef As Variant
    For Each ef In entityFolders
        IndexTBYear dict, CStr(ef), "TB 24", "tb24"
        IndexTBYear dict, CStr(ef), "TB 25", "tb25"
    Next ef

    Set BuildTBIndex = dict
End Function

' Helper for BuildTBIndex: scans one year subfolder of one entity.
Private Sub IndexTBYear(dict As Object, entityFolder As String, _
                        subName As String, yearKey As String)
    Dim subPath As String : subPath = entityFolder & "\" & subName
    If Dir(subPath, vbDirectory) = "" Then Exit Sub

    ' Collect filenames BEFORE opening any workbook (Dir is not reentrant)
    Dim files As New Collection
    AddFilesByPattern files, subPath, "*.xlsx"
    AddFilesByPattern files, subPath, "*.xlsm"
    AddFilesByPattern files, subPath, "*.xls"

    Dim fn As Variant
    For Each fn In files
        Dim wb As Workbook
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(subPath & "\" & CStr(fn), _
                                ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        If Not wb Is Nothing Then
            Dim a1 As String
            a1 = Trim(CStr(wb.Sheets(1).Cells(1, 1).Value))
            wb.Close SaveChanges:=False
            Set wb = Nothing
            If a1 <> "" Then
                If Not dict.Exists(a1) Then
                    Dim rec As Object
                    Set rec = CreateObject("Scripting.Dictionary")
                    rec.CompareMode = 1
                    rec("folder") = entityFolder
                    rec("tb24") = ""
                    rec("tb25") = ""
                    dict.Add a1, rec
                End If
                Dim existing As Object
                Set existing = dict(a1)
                existing(yearKey) = CStr(fn)
                ' Keep the first-seen folder; only overwrite if it was empty
                If CStr(existing("folder")) = "" Then
                    existing("folder") = entityFolder
                End If
            End If
        End If
    Next fn
End Sub

' =============================================================
' LoadConsolidatedGL
' Opens the consolidated GL workbook once and reads the used
' range of sheet 1 into a 1-based 2D Variant array via
' Range.Value.
'
' If the workbook is already open in the current Excel instance
' the existing Workbook object is reused and NOT closed.
'
' Output:
'   glData - 2D Variant array (rows x cols); Empty on failure.
'   errMsg - "" on success, otherwise a human-readable reason.
'            (glData stays Empty when errMsg is non-empty so the
'            caller can fall back gracefully per row.)
' =============================================================
Private Sub LoadConsolidatedGL(path As String, ByRef glData As Variant, _
                               ByRef errMsg As String)
    glData = Empty
    errMsg = ""

    Dim nameOnly As String : nameOnly = Mid(path, InStrRev(path, "\") + 1)
    Dim wb As Workbook
    Dim alreadyOpen As Boolean

    On Error Resume Next
    Set wb = Workbooks(nameOnly)
    On Error GoTo 0
    alreadyOpen = Not (wb Is Nothing)

    If Not alreadyOpen Then
        On Error Resume Next
        Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        If wb Is Nothing Then
            errMsg = "could not open consolidated GL"
            Exit Sub
        End If
    End If

    Dim ws As Worksheet : Set ws = wb.Sheets(1)
    Dim lastRow As Long : lastRow = LastUsedRow(ws)
    Dim lastCol As Long : lastCol = LastUsedCol(ws)
    If lastRow < 1 Or lastCol < 8 Then
        errMsg = "consolidated GL appears empty or has fewer than 8 columns"
        If Not alreadyOpen Then wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Range.Value on a single cell returns a scalar, not a 2D array,
    ' so promote that edge case manually.
    If lastRow = 1 And lastCol = 1 Then
        ReDim glData(1 To 1, 1 To 1)
        glData(1, 1) = ws.Cells(1, 1).Value
    Else
        glData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    End If

    If Not alreadyOpen Then wb.Close SaveChanges:=False
End Sub

' =============================================================
' FilterConsolidatedGL
' Writes a new workbook at outputPath containing:
'   - Row 1 from glData (header, always kept).
'   - Every subsequent row whose column H (index 8) value equals
'     one of the non-empty mapping GL names (Trim +
'     case-insensitive exact equality).
'
' Returns the number of data rows kept (header excluded).
'
' Uses a single .Value = 2D-Variant assignment on the destination
' range to keep the write fast and avoid clipboard side-effects.
' =============================================================
Private Function FilterConsolidatedGL(ByRef glData As Variant, _
                                      glNames As Variant, _
                                      outputPath As String) As Long
    FilterConsolidatedGL = 0

    ' Mapping GL names -> case-insensitive lookup set
    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")
    nameDict.CompareMode = 1 ' vbTextCompare

    Dim k As Long
    For k = LBound(glNames) To UBound(glNames)
        Dim nm As String : nm = Trim(CStr(glNames(k)))
        If nm <> "" Then
            If Not nameDict.Exists(nm) Then nameDict.Add nm, True
        End If
    Next k
    If nameDict.Count = 0 Then Exit Function

    Dim nRows As Long : nRows = UBound(glData, 1)
    Dim nCols As Long : nCols = UBound(glData, 2)
    If nRows < 1 Or nCols < 8 Then Exit Function

    ' Count matches first so we can size the output array exactly.
    Dim keep As Long : keep = 0
    Dim i As Long
    For i = 2 To nRows
        Dim h As String : h = Trim(CStr(glData(i, 8)))
        If h <> "" Then
            If nameDict.Exists(h) Then keep = keep + 1
        End If
    Next i

    ' Build output: header row + kept rows (header always kept,
    ' even when no data rows matched).
    Dim outRows As Long : outRows = keep + 1
    Dim outArr() As Variant
    ReDim outArr(1 To outRows, 1 To nCols)

    Dim c As Long
    For c = 1 To nCols
        outArr(1, c) = glData(1, c)
    Next c

    Dim dstRow As Long : dstRow = 1
    For i = 2 To nRows
        Dim h2 As String : h2 = Trim(CStr(glData(i, 8)))
        If h2 <> "" Then
            If nameDict.Exists(h2) Then
                dstRow = dstRow + 1
                For c = 1 To nCols
                    outArr(dstRow, c) = glData(i, c)
                Next c
            End If
        End If
    Next i

    ' Write to a brand-new workbook and SaveAs .xlsx.
    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    Dim newWs As Worksheet : Set newWs = newWb.Sheets(1)
    newWs.Range(newWs.Cells(1, 1), newWs.Cells(outRows, nCols)).Value = outArr

    Application.DisplayAlerts = False
    newWb.SaveAs outputPath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
    Set newWb = Nothing

    FilterConsolidatedGL = keep
End Function
