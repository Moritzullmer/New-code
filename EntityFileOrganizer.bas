Attribute VB_Name = "EntityFileOrganizer"
Option Explicit

' =============================================================
' ENTITY FILE ORGANIZER
'
' Reads a mapping workbook listing entities with TB codes and GL
' names, then scans a source root whose immediate subfolders are
' entity folders (each with TB 24/, TB 25/, GL/ subfolders).  For
' every mapping row it produces an output folder named after the
' entity, containing:
'   - the matched TB 24 file (copied unchanged)
'   - the matched TB 25 file (copied unchanged)
'   - GL.xlsx = merge of every matched GL source (headers from the
'     first file, all subsequent files' rows 2..last appended).
'
' Mapping layout on sheet 1 of the mapping workbook:
'   Row 3 headers:   A=Entity Name  B=TB #   C=GL 1  D=GL 2  E=GL 3 ...
'   Row 4+ data:     one row per entity
'   GL columns:      read until the row-3 header cell is empty
'
' Matching rules:
'   - TB file:  cell A1 equals the row's TB # (Trim, case-insensitive).
'   - GL file:  cell H2 equals the row's GL name (Trim, case-insensitive).
' =============================================================

' =============================================================
' PUBLIC ENTRY POINT
' Prompts for mapping workbook, source root, and output root,
' then builds one output subfolder per mapping row.
' =============================================================
Public Sub Main()

    ' ---- Pickers ----
    Dim mappingPath As String
    mappingPath = PickFile("Select the mapping workbook")
    If mappingPath = "" Then Exit Sub

    Dim sourceRoot As String
    sourceRoot = PickFolder("Select the source root folder (contains entity subfolders)")
    If sourceRoot = "" Then Exit Sub

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

        ' --- TB 24 copy ---
        If tb24File = "" Then
            rowWarn = rowWarn & "TB 24 missing; "
        Else
            Dim tb24Dst As String : tb24Dst = outFolder & "\" & tb24File
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
            Dim tb25Dst As String : tb25Dst = outFolder & "\" & tb25File
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

        ' --- GL merge ---
        Dim mergedPath As String : mergedPath = outFolder & "\GL.xlsx"
        On Error Resume Next
        If Dir(mergedPath) <> "" Then Kill mergedPath
        On Error GoTo Cleanup

        Dim expectedGL As Long : expectedGL = CountNonEmpty(glNames)
        Dim mergedCount As Long : mergedCount = 0
        On Error Resume Next
        mergedCount = MergeGLFiles(entityFolder & "\GL", glNames, mergedPath)
        If Err.Number <> 0 Then
            rowWarn = rowWarn & "GL merge failed (" & Err.Description & "); "
            Err.Clear
        End If
        On Error GoTo Cleanup
        If expectedGL = 0 Then
            rowWarn = rowWarn & "no GL names in mapping; "
        ElseIf mergedCount = 0 Then
            rowWarn = rowWarn & "no GL files matched H2 in " & entityFolder & "\GL; "
        ElseIf mergedCount < expectedGL Then
            rowWarn = rowWarn & "only " & mergedCount & " of " & _
                      expectedGL & " GL files matched; "
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
' MergeAppend
' Appends rows 2..srcLastRow of wsSource below the current last
' row of wsTarget, across columns 1..srcLastCol.
'
' Uses a single .Value = .Value assignment, which is much faster
' than Copy/PasteSpecial and avoids clipboard side-effects.
' =============================================================
Private Sub MergeAppend(wsTarget As Worksheet, wsSource As Worksheet)
    Dim srcLastRow As Long : srcLastRow = LastUsedRow(wsSource)
    Dim srcLastCol As Long : srcLastCol = LastUsedCol(wsSource)
    If srcLastRow < 2 Or srcLastCol < 1 Then Exit Sub

    Dim tgtLastRow As Long : tgtLastRow = LastUsedRow(wsTarget)
    Dim dstStart As Long : dstStart = tgtLastRow + 1
    Dim nRows As Long : nRows = srcLastRow - 1 ' rows 2..srcLastRow inclusive

    Dim srcRange As Range
    Set srcRange = wsSource.Range(wsSource.Cells(2, 1), _
                                  wsSource.Cells(srcLastRow, srcLastCol))

    Dim dstRange As Range
    Set dstRange = wsTarget.Range(wsTarget.Cells(dstStart, 1), _
                                  wsTarget.Cells(dstStart + nRows - 1, srcLastCol))

    dstRange.Value = srcRange.Value
End Sub

' =============================================================
' MergeGLFiles
' Resolves each mapping GL name (via cell H2 = name) to a file in
' entityGLFolder, then writes a merged workbook to outputPath:
'   - First matched file becomes the base (header row preserved).
'   - Subsequent matched files: rows 2..last appended in order.
'
' Returns the number of GL source files merged (0 = none).
'
' The source GL files are opened ReadOnly and never saved back;
' the merged workbook is written via SaveAs xlOpenXMLWorkbook so
' the result is always a genuine .xlsx regardless of source
' extension (.xls / .xlsm / .xlsx).
' =============================================================
Private Function MergeGLFiles(entityGLFolder As String, _
                              glNames As Variant, _
                              outputPath As String) As Long
    MergeGLFiles = 0
    If Dir(entityGLFolder, vbDirectory) = "" Then Exit Function

    ' ---- Collect every GL file in the folder (Dir is not reentrant) ----
    Dim glFiles As New Collection
    AddFilesByPattern glFiles, entityGLFolder, "*.xlsx"
    AddFilesByPattern glFiles, entityGLFolder, "*.xlsm"
    AddFilesByPattern glFiles, entityGLFolder, "*.xls"
    If glFiles.Count = 0 Then Exit Function

    ' ---- Build H2-value -> filename index ----
    Dim h2map As Object
    Set h2map = CreateObject("Scripting.Dictionary")
    h2map.CompareMode = 1 ' vbTextCompare

    Dim fn As Variant
    For Each fn In glFiles
        Dim wb As Workbook
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(entityGLFolder & "\" & CStr(fn), _
                                ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        If Not wb Is Nothing Then
            Dim h2 As String
            h2 = Trim(CStr(wb.Sheets(1).Cells(2, 8).Value))
            wb.Close SaveChanges:=False
            Set wb = Nothing
            If h2 <> "" And Not h2map.Exists(h2) Then h2map(h2) = CStr(fn)
        End If
    Next fn

    ' ---- Resolve each mapping GL name, preserving mapping order ----
    Dim resolved As New Collection
    Dim i As Long
    For i = LBound(glNames) To UBound(glNames)
        Dim gname As String : gname = Trim(CStr(glNames(i)))
        If gname <> "" Then
            If h2map.Exists(gname) Then resolved.Add CStr(h2map(gname))
        End If
    Next i
    If resolved.Count = 0 Then Exit Function

    ' ---- Open first match, append the rest, SaveAs output as .xlsx ----
    Dim outWb As Workbook
    Set outWb = Nothing
    On Error Resume Next
    Set outWb = Workbooks.Open(entityGLFolder & "\" & CStr(resolved(1)), _
                               ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    If outWb Is Nothing Then Exit Function
    Dim outWs As Worksheet : Set outWs = outWb.Sheets(1)

    Dim j As Long
    For j = 2 To resolved.Count
        Dim srcWb As Workbook
        Set srcWb = Nothing
        On Error Resume Next
        Set srcWb = Workbooks.Open(entityGLFolder & "\" & CStr(resolved(j)), _
                                   ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        If Not srcWb Is Nothing Then
            MergeAppend outWs, srcWb.Sheets(1)
            srcWb.Close SaveChanges:=False
            Set srcWb = Nothing
        End If
    Next j

    Application.DisplayAlerts = False
    outWb.SaveAs outputPath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    outWb.Close SaveChanges:=False
    Set outWb = Nothing

    MergeGLFiles = resolved.Count
End Function
