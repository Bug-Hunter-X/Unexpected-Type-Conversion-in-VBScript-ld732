Function f(a)
  If IsNumeric(a) Then
    a = CInt(a) 'Explicitly convert to integer
    If a = 1 Then
      f = 1
    ElseIf a = 2 Then
      f = 2
    ElseIf a = 3 Then
      f = 3
    Else
      f = 4
    End If
  Else
    f = -1 'Handle non-numeric input
  End If
End Function
MsgBox f(1)
MsgBox f("1")
MsgBox f(4)