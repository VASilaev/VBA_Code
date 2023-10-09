Public Function SelectOneRow(byval sql As String, ParamArray params()) As Variant
  Dim rsdao, objField
  if UBound(params) >= 0 then 
    sql = printf(sql, params)
  end if    
  Set SelectOneRow = CreateObject("Scripting.Dictionary")
  SelectOneRow.CompareMode = 1 ' 1 = TextCompare
  Set rsdao = CurrentProject.Connection.Execute(sql)
  If Not rsdao.EOF Then
   For Each objField In rsdao.Fields
    SelectOneRow.Add objField.Name, objField.Value
   Next
  End If
  Set objField = Nothing
  rsdao.Close
  Set rsdao = Nothing
End Function

Public Function SelectOneValue(byval sql As String, ParamArray params()) As Variant
  Dim rsdao
  if UBound(params) >= 0 then 
    sql = printf(sql, params)
  end if          
  Set rsdao = CurrentProject.Connection.Execute(sql)
  On Error GoTo noRecord
  If rsdao.EOF Then SelectOneValue = Empty Else SelectOneValue = rsdao.Fields(0).Value
  rsdao.Close
  Set rsdao = Nothing
  Exit Function
noRecord:
  SelectOneValue = Empty
  rsdao.Close
  Set rsdao = Nothing
End Function

Public Function SelectFirstColumn(sql As String, ParamArray params()) As Variant()
  Dim nvCnt As Long, rsdao
  Dim avResult() As Variant
  ReDim avResult(256)
  if UBound(params) >= 0 then 
    sql = printf(sql, params)
  end if                   
  Set rsdao = CurrentProject.Connection.Execute(sql)
  nvCnt = 0
  Do While Not rsdao.EOF
    If nvCnt > UBound(avResult) Then ReDim Preserve avResult(UBound(avResult) + 256)
    avResult(nvCnt) = rsdao.Fields(0).Value
    nvCnt = nvCnt + 1
    rsdao.MoveNext
  Loop
  If nvCnt = 0 Then
    ReDim avResult(0)
    avResult(0) = Empty
  Else
    ReDim Preserve avResult(nvCnt - 1)
  End If
  rsdao.Close
  Set rsdao = Nothing
  SelectFirstColumn = avResult
End Function

Public Function OpenDataSet(byval sql As String, ParamArray params()) As Variant
  Dim rsdao, objField
  if UBound(params) >= 0 then 
    sql = printf(sql, params)
  end if    
  Set OpenDataSet = CurrentProject.Connection.Execute(sql)
End Function
