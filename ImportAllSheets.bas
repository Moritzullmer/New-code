Sub ImportAllSheets()
    Dim srcPath As String
    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim ws As Worksheet
    Dim existingNames As Object
    Dim newName As String
    Dim counter As Integer

    ' Prompt user to select the source file
    srcPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        Title:="Select Source Workbook")

    If srcPath = "False" Then
        MsgBox "No file selected. Operation cancelled.", vbInformation
        Exit Sub
    End If

    Set dstWb = ThisWorkbook

    ' Build a set of existing sheet names for collision detection
    Set existingNames = CreateObject("Scripting.Dictionary")
    Dim existing As Worksheet
    For Each existing In dstWb.Worksheets
        existingNames(LCase(existing.Name)) = True
    Next existing

    Application.ScreenUpdating = False

    Set srcWb = Workbooks.Open(srcPath, ReadOnly:=True)

    For Each ws In srcWb.Worksheets
        newName = ws.Name
        ' Resolve name collisions by appending _2, _3, etc.
        If existingNames.Exists(LCase(newName)) Then
            counter = 2
            Do While existingNames.Exists(LCase(newName & "_" & counter))
                counter = counter + 1
            Loop
            newName = newName & "_" & counter
        End If

        ws.Copy After:=dstWb.Sheets(dstWb.Sheets.Count)
        dstWb.Sheets(dstWb.Sheets.Count).Name = newName
        existingNames(LCase(newName)) = True
    Next ws

    srcWb.Close SaveChanges:=False

    Application.ScreenUpdating = True

    MsgBox "Done! " & srcWb.Sheets.Count & " sheet(s) imported from:" & vbLf & srcPath, vbInformation
End Sub
