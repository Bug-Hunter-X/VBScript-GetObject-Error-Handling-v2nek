Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

' Example usage
Set myExcel = GetObject("C:\\path\\to\\your\\excel.xls")
if myExcel is nothing then
    msgbox "Excel file not found."
exit sub
end if