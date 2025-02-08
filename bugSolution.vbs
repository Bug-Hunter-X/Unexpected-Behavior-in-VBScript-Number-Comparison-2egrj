Function f(a,b)
  On Error Resume Next
  If IsNumeric(a) And IsNumeric(b) Then
    If a > b Then
      MsgBox "a is greater than b"
    ElseIf a < b Then
      MsgBox "a is less than b"
    Else
      MsgBox "a is equal to b"
    End If
  Else
    MsgBox "Invalid Input: Please enter numbers only."
  End If
  On Error GoTo 0
end function

Function main()
  Dim a,b
  a = InputBox("Enter the value of a")
  b = InputBox("Enter the value of b")
  f(a,b)
end function
main()