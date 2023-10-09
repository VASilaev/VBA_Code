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

Public Function ToSQL(Val)
    Select Case VarType(Val)
      Case 8
        ToSQL = StrToSQL(Val)
      Case 7
        ToSQL = DateToSQL(Val)
      Case 0,1
        ToSQL = "NULL"
      Case 11
        If Val Then ToSQL = "true" Else ToSQL = "false"
      Case Else
        ToSQL = Val & ""
      End Select
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
