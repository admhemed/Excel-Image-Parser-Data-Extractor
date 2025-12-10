Sub PastePictureAndSave()
    Dim ws As Worksheet
    Dim tgtCell As Range
    Dim shp As Shape
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim chObj As ChartObject
    Dim guid As String
    Dim currentRow As Long
    Dim idCol As Long
    Dim fileCol As Long
    
    Set ws = ActiveSheet
    Set tgtCell = ActiveCell   ' target cell where the user wants to attach the picture
    
    ' Paste (image from clipboard)
    On Error Resume Next
    ws.Paste
    On Error GoTo 0
    
    If ws.Shapes.Count = 0 Then
        ' No picture pasted, silently exit
        Exit Sub
    End If
    
    ' Assume the last shape is the newly pasted picture
    Set shp = ws.Shapes(ws.Shapes.Count)
    
    ' Position the picture over the target cell (before resizing)
    shp.Left = tgtCell.Left
    shp.Top = tgtCell.Top
    
    ' Create images folder next to this workbook
    folderPath = ThisWorkbook.Path & "\images\"
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Generate GUID-based file name (without braces) and use .jpg
    guid = CreateObject("Scriptlet.TypeLib").guid
    guid = Mid$(guid, 2, 36)          ' remove { and }
    fileName = guid & ".jpg"
    fullPath = folderPath & fileName
    
    ' --- Export via temporary chart (at original size) ---
    shp.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    
    Set chObj = ws.ChartObjects.Add(Left:=shp.Left, Top:=shp.Top, _
                                    Width:=shp.Width, Height:=shp.Height)
    
    chObj.Activate
    chObj.Chart.Paste
    DoEvents  ' let Excel render the picture
    
    chObj.Chart.Export fileName:=fullPath, FilterName:="JPG"
    chObj.Delete
    ' --- End of export logic ---
    
    ' --- Resize the picture on the sheet to 60x60 ---
    With shp
        .LockAspectRatio = msoFalse   ' set msoTrue if you want to keep aspect ratio
        .Width = 60
        .Height = 60
        .Left = tgtCell.Left
        .Top = tgtCell.Top
    End With
    
    ' Columns for ID (left) and file name (right)
    idCol = tgtCell.Column - 1
    fileCol = tgtCell.Column + 1
    
    ' Write on the same row: left = GUID (no extension), right = file name with extension
    ws.Cells(tgtCell.Row, idCol).Value = guid
    ws.Cells(tgtCell.Row, fileCol).Value = fileName
    
    ' --- Repeat GUID and file name on subsequent rows until an empty row (first 6 columns empty) ---
    currentRow = tgtCell.Row + 1
    
    Do While Application.WorksheetFunction.CountA(ws.Range(ws.Cells(currentRow, 1), _
                                                           ws.Cells(currentRow, 6))) > 0
        ws.Cells(currentRow, idCol).Value = guid
        ws.Cells(currentRow, fileCol).Value = fileName
        currentRow = currentRow + 1
    Loop
End Sub
