' Excel Macro
' Import Multiple CSV files into separate tabs

Sub CombineCsvFiles()
    Dim xFilesToOpen As Variant
    Dim I As Integer
    Dim xWb As Workbook
    Dim xTempWb As Workbook
    Dim xDelimiter As String
    Dim xScreen As Boolean
    On Error GoTo ErrHandler
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    xDelimiter = "|"
    xFilesToOpen = Application.GetOpenFilename("Text Files (*.csv), *.csv", , "Combine CSV Files", , True)
    If TypeName(xFilesToOpen) = "Boolean" Then
        MsgBox "No files were selected", , "Combine CSV Files"
        GoTo ExitHandler
    End If
    I = 1
    Set xTempWb = Workbooks.Open(xFilesToOpen(I))
    xTempWb.Sheets(1).Copy
    Set xWb = Application.ActiveWorkbook
    xTempWb.Close False
    Do While I < UBound(xFilesToOpen)
        I = I + 1
        Set xTempWb = Workbooks.Open(xFilesToOpen(I))
        xTempWb.Sheets(1).Move , xWb.Sheets(xWb.Sheets.Count)
    Loop
ExitHandler:
    Application.ScreenUpdating = xScreen
    Set xWb = Nothing
    Set xTempWb = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, , "Combine CSV Files"
    Resume ExitHandler
End Sub

