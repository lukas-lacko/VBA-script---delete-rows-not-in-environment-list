Sub DeleteRowsNotContainingDefinedEnvIDs()
    Dim ws As Worksheet
    Dim colNum As Long
    Dim lastRow As Long
    Dim i As Long
    Dim targetEnvIDs As Variant
    Dim keepHeader As Boolean

    ' === Configuration ===
    colNum = 2 ' Column B (change to your desired column)
    targetEnvIDs = Array("EnvID1", "EnvID2", "EnvID3")  ' Env IDs to match
    keepHeader = True ' Set to False if you want to delete the header too
    ' ======================

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    Application.ScreenUpdating = False

    ' Loop from bottom to top to avoid skipping rows after deletion
    For i = lastRow To IIf(keepHeader, 2, 1) Step -1
        Dim cellValue As String
        cellValue = LCase(ws.Cells(i, colNum).Value)
        Dim matchFound As Boolean
        matchFound = False

        Dim s As Variant
        For Each s In targetEnvIDs
            If InStr(1, cellValue, LCase(s), vbTextCompare) > 0 Then
                matchFound = True
                Exit For
            End If
        Next s

        If Not matchFound Then
            ws.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Rows deleted successfully!", vbInformation
End Sub
