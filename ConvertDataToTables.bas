' Excel Macro
' Convert all Excel Sheets to Tables

Sub ConvertDataToTables()
'
' ConvertDataToTables Macro
' Make all sheets tables
'
  ActiveWorkbook.DefaultTableStyle = "TableStyleLight1"
  Sheets(1).Select

  For I = 1 To Sheets.Count
    Sheets(I).Activate
  
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    If ActiveSheet.ListObjects.Count < 1 Then
      ActiveSheet.ListObjects.Add.Name = ActiveSheet.Name
    End If
    
  Next I

End Sub



