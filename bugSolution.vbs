Function GetObjectSafe(path)
  Dim obj, ErrNum
  On Error Resume Next
  Set obj = GetObject(path)
  ErrNum = Err.Number
  On Error GoTo 0
  If ErrNum <> 0 Then
    WScript.Echo "Error: Could not open file. Error Number: " & ErrNum & ". Check the file path and make sure the file exists and you have permissions to open it." 
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

' Example usage
Set myExcel = GetObjectSafe("C:\\path\\to\\your\\excel.xls")
If myExcel Is Nothing Then
    WScript.Quit
Else
  'Do something with myExcel
End If