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

Public Function ToSQL(pValue)
'На основе типа данных преобразует значение в SQL литерал
'#param pValue - Значение для преобразования
  Select Case VarType(pValue)
    Case vbString
      ToSQL = "'" & Replace(pValue, "'", "''") & "'"
    Case vbDate
      If pValue = CLng(pValue) Then
        ToSQL = "#" & Format(pValue, "mm\/dd\/yyyy") & "#"
      ElseIf pValue < 1 Then
        ToSQL = "#" & Format(pValue, "hh:nn:ss") & "#"
      Else
        ToSQL = "#" & Format(pValue, "mm\/dd\/yyyy hh:nn:ss") & "#"
      End If
    Case vbEmpty, vbNull
      ToSQL = "NULL"
    Case vbBoolean
      If pValue Then ToSQL = "true" Else ToSQL = "false"
    Case vbInteger, vbLong, 20
      ToSQL = pValue & ""
    Case vbSingle,vbDouble,vbCurrency, vbDecimal
      ToSQL = Replace(pValue & "", ",", ".")
    'vbByte ?? char
    Case Else
      if isArray(pValue) then 
        dim vElement
        ToSQL = ""
        for each vElement in pValue
          if len(ToSQL) = 0 then 
            ToSQL = ToSQL(vElement) 
          else 
            ToSQL = ToSQL & ", " & ToSQL(vElement)  
          end if
        next
        if ToSQL ="" then ToSQL = "NULL"
      else
        err.rise 1001,,"Unsupported type of SQL value!"
      end if
  End Select
End Function

Public Function printf(ByVal mask As String, ParamArray tokens()) As String
'Вставляет в текст значения параметров. Параметр задается своим номером заключенным в фигурные скобки
'#param mask - Строка для форматирования. 
'  Место вставки укажите в формате {0}, {1}... где число - это индекс передаваемого параметра
'  Если указать в формате {0;SQL}, {1;SQL}... то значение будет отформатировано в SQL литерал
'#param tokens - список значений для подстановки. Если единственным параметром передать массив, то он будет развернут. 
'  Для Предотвращения данного поведения, передйте дополнительный параметр 
  Dim i As Long, T
  T = tokens
  If UBound(T) = 0 Then If IsArray(T(0)) Then T = T(0)
  For i = 0 To UBound(T)
      mask = Replace$(mask, "{" & i & "}", T(i))
  Next
  If InStr(mask, "{") > 0 Then
    With CreateObject("VBScript.RegExp")
      .Global = True
      .pattern = "\{(\d+)\s*;([^\}]*)\}"
      Dim token
      For Each token In .Execute(mask)
        if UCase(token.subMatches(1)) = "SQL" then 
          mask = Replace$(mask, token, ToSQL(T(CInt(token.subMatches(0)))))
        else
          mask = Replace$(mask, token, Format(T(CInt(token.subMatches(0))), token.subMatches(1)))
        end if
      Next
    End With
  End If
  printf = mask
End Function

Public Function MaxDate(Optional vvDefault = Null) As Date
 If IsNull(vvDefault) Then
  MaxDate = #1/1/2100#
 Else
  MaxDate = vvDefault
 End If
End Function

Public Function MinDate(Optional vvDefault = Null) As Date
 If IsNull(vvDefault) Then
  MinDate = #1/1/1900#
 Else
  MinDate = vvDefault
 End If
End Function

'Use NZ
Public Function NVL(condition, Value)
 If IsNull(condition) Then NVL = Value Else NVL = condition
End Function        

Public Function bitAnd(o1 As Long, o2 As Long) As Long
 bitAnd = o1 And o2
End Function          
