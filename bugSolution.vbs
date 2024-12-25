Function f(a, b)
  If IsEmpty(a) Or IsNull(a) Then
    a = 0
  ElseIf Not IsNumeric(a) Then
    Err.Raise 13, , "Type mismatch: a must be numeric" 
  End If
  If IsEmpty(b) Or IsNull(b) Then
    b = 0
  ElseIf Not IsNumeric(b) Then
    Err.Raise 13, , "Type mismatch: b must be numeric" 
  End If
  c = a + b
  f = c
End Function
MsgBox f(1, Empty)