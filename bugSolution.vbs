Function MyFunction(param)
  If IsEmpty(param) Or param = "" Then
    Err.Raise vbError, , "Parameter cannot be empty or null"
  Else
    ' ... rest of the function
  End If
End Function

' Example usage:
On Error GoTo ErrorHandler

Dim result
result = MyFunction(Null)  'This will generate the error 
result = MyFunction("")    'This will generate the error
result = MyFunction("Hello") ' This will execute normally

Exit Sub

ErrorHandler:
  MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
  Err.Clear
End Sub