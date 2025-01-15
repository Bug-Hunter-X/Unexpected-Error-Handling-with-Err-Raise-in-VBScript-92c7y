Function MyFunction(param)
  If IsEmpty(param) Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
End Function