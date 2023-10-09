Option Compare Database
Option Explicit

Public Function DateToSQL(pDate)
  If pDate = CLng(pDate) Then
    DateToSQL = "#" & Format(pDate, "mm\/dd\/yyyy") & "#"
  ElseIf pDate < 1 Then
    DateToSQL = "#" & Format(pDate, "hh:nn:ss") & "#"
  Else
    DateToSQL = "#" & Format(pDate, "mm\/dd\/yyyy hh:nn:ss") & "#"
  End If
End Function
 
Public Function StrToSQL(pText)
  StrToSQL = "'" & Replace(pText, "'", "''") & "'"
End Function

Public Function SelectOneRow(sql As String) As Variant
 Dim rsdao, objField
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

Public Function ArrayToInStatement(apValues)
  If IsArray(apValues) Then
    If UBound(apValues) >= 0 Then
      Dim Val
      ArrayToInStatement = ""
      For Each Val In apValues
        If ArrayToInStatement > "" Then ArrayToInStatement = ArrayToInStatement & ", "
        Select Case VarType(Val)
          Case 8
            ArrayToInStatement = ArrayToInStatement & StrToSQL(Val)
          Case 7
            ArrayToInStatement = ArrayToInStatement & DateToSQL(Val)
          Case 0, 1
            'Больше нужно для отладки
            ArrayToInStatement = ArrayToInStatement & "Null"

          Case Else
            'Числа записываются в естественном виде
            ArrayToInStatement = ArrayToInStatement & Val
          End Select
      Next
      ArrayToInStatement = "(" & ArrayToInStatement & ")"
    Else
      ArrayToInStatement = "(Null)"
    End If
  Else
    ArrayToInStatement = "(Null)"
  End If
End Function

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

Public Function printf(ByVal mask As String, ParamArray tokens()) As String
    Dim i As Long, T
    T = tokens
    If UBound(T) = 0 Then If IsArray(T(0)) Then T = T(0)
    For i = 0 To UBound(T)
        mask = Replace$(mask, "{" & i & "}", T(i))
    Next
    If InStr(mask, "{") > 0 Then
      'Дополнительная обработка по маске
      With CreateObject("VBScript.RegExp")
        .Global = True
        .pattern = "\{(\d+)\s*;([^\}]*)\}"
        Dim token
        For Each token In .Execute(mask)
          mask = Replace$(mask, token, Format(T(CInt(token.subMatches(0))), token.subMatches(1)))
        Next
      End With
    End If
    printf = mask
End Function

Public Function getNumber(sMethodNumeration, ParamArray apSection())
  Dim aSection, Method, Row
  aSection = apSection
  
  Set Method = SelectOneRow("select id, trim(sFormat) as sFormat from tNumerationMethod where UCase(sName) = " & StrToSQL(UCase(sMethodNumeration)))
  
  If Method.Count = 0 Then
    Err.Raise 1000, , printf("Не удалось найти метод нумерации с кодом [{0}]", sMethodNumeration)
  Else
    DBEngine.Workspaces(0).BeginTrans
    On Error GoTo Rollback
    Set Row = SelectOneRow(printf("select id, nLastNumber from tNumeration as t where t.idNumerationMethod = {0} and t.sSection = {1}", Method("id"), StrToSQL(ArrayToLiteral(aSection))))
    If Row.Count = 0 Then
      getNumber = Replace(printf(Method("sFormat"), aSection), "{NUM}", 1)
      CurrentDb().Execute printf("insert into tNumeration(idNumerationMethod,nLastNumber,sSection) values ({0}, 1, {1})", Method("id"), StrToSQL(ArrayToLiteral(aSection)))
    Else
      getNumber = Replace(printf(Method("sFormat"), aSection), "{NUM}", Row("nLastNumber") + 1)
      CurrentDb().Execute printf("update tNumeration set nLastNumber = {1} where id = {0}", Row("id"), Row("nLastNumber") + 1)
    End If
    DBEngine.Workspaces(0).CommitTrans
  End If
  Exit Function
Rollback:
   On Error GoTo 0
   DBEngine.Workspaces(0).Rollback
End Function

'Создает необходимые таблицы
Public Sub InstallNumeratiomSystem()
  Dim vTbl, vFld, vDB
  
  Set vDB = CurrentDb()
  On Error GoTo onCreate
  Set vTbl = vDB.TableDefs("tNumeration")
  On Error GoTo 0
  Exit Sub
onCreate:
  On Error GoTo 0
  
  vDB.Execute "CREATE TABLE tNumerationMethod (id counter CONSTRAINT PK_NumerationMethod PRIMARY KEY, sName CHAR(255), sFormat CHAR(255));"
    
  vDB.Execute "CREATE TABLE tNumeration (id counter CONSTRAINT PK_Numeration PRIMARY KEY, idNumerationMethod long, nLastNumber long, sSection CHAR(255), CONSTRAINT FK_tNumeration_NumerationMethod FOREIGN KEY (idNumerationMethod) REFERENCES tNumerationMethod(id));"
    
  vDB.TableDefs.Refresh

End Sub


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

'Рекурсивный поиск
Public Function FindAllAnalog(spPartNum, Optional bpDebug = False)
 
  Dim aAnalog, rsdao, SQL, sDebug
  
  aAnalog = Array(spPartNum)
  
  Do While UBound(aAnalog) >= 0
    FindAllAnalog = ConcatArray(FindAllAnalog, aAnalog)
    
    'здесь нужно заменить запрос который вернет новый уровень иерархии без повторений
    SQL = printf("select PN from (" & _
            "select [PN замена] as PN from замены_аналоги where PN in {0} " & _
            "union select PN from замены_аналоги where [PN замена] in {0} " & _
          ") where not PN in {1}", ArrayToInStatement(aAnalog), ArrayToInStatement(FindAllAnalog))
    
    aAnalog = Array()
    
    Set rsdao = CurrentProject.Connection.Execute(SQL)
    
    Do While Not rsdao.EOF
      addInArray aAnalog, rsdao.Fields("PN").Value
      rsdao.MoveNext
    Loop
    
    If bpDebug And UBound(aAnalog) >= 0 Then addInArray FindAllAnalog, "-> "
    
    rsdao.Close
    Set rsdao = Nothing
  Loop
  
  FindAllAnalog = Filter(FindAllAnalog, spPartNum, False)
  
End Function
