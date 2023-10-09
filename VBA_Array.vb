Public Function ArrayToInStatement(apValues)
  If IsArray(apValues) Then
    If UBound(apValues) >= 0 Then
      Dim Val
      ArrayToInStatement = ""
      For Each Val In apValues
        If ArrayToInStatement > "" Then ArrayToInStatement = ArrayToInStatement & ", "
        ArrayToInStatement = ArrayToInStatement & ToSQL(Val)
      Next
      ArrayToInStatement = "(" & ArrayToInStatement & ")"
    Else
      ArrayToInStatement = "(Null)"
    End If
  Else
    ArrayToInStatement = "(Null)"
  End If
End Function
        
Public Function ConcatArray(aData1, aData2)
  Dim tmpArray, i, j
  tmpArray = aData1
  If IsEmpty(tmpArray) Then tmpArray = Array()
  If Not IsArray(tmpArray) Then tmpArray = Array(tmpArray)
  If IsArray(aData2) Then
    i = UBound(tmpArray) + 1
    ReDim Preserve tmpArray(UBound(tmpArray) + UBound(aData2) + 1)
    For j = 0 To UBound(aData2)
      tmpArray(i + j) = aData2(j)
    Next
  End If
  ConcatArray = tmpArray
  tmpArray = Empty
End Function

Public Sub addInArray(ByRef spArray, ByRef pItem)
  If IsEmpty(spArray) Then
    spArray = Array(pItem)
  Else
    ReDim Preserve spArray(UBound(spArray) + 1)
    spArray(UBound(spArray)) = pItem
  End If
End Sub

Function ArrayToLiteral(spArr)
  Dim el
  ArrayToLiteral = ""
  For Each el In spArr
    Select Case VarType(el)
      Case 8
        el = StrToSQL(el)
      Case 7
        el = DateToSQL(el)
      Case 0
        el = "E"
      Case 1
        el = "N"
      Case 11
        If el Then el = "t" Else el = "f"
      Case Else
        el = el & ""
      End Select
    If Len(ArrayToLiteral) > 0 Then el = "," & el
    ArrayToLiteral = ArrayToLiteral & el
  Next
End Function

