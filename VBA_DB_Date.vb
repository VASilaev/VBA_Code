Public Function MaxDate(Optional dDate = Null) As Date
'Возвращает 01.01.2100 если значение является Null, обычно используется при проверке открытых диапазонов
'#param dDate: Значение, которое проверяется на Null значение
 If IsNull(dDate) Then
  MaxDate = #1/1/2100#
 Else
  MaxDate = dDate
 End If
End Function

Public Function MinDate(Optional dDate = Null) As Date
'Возвращает 01.01.1900 если значение является Null, обычно используется при проверке открытых диапазонов
'#param dDate: Значение, которое проверяется на Null значение
 If IsNull(dDate) Then
  MinDate = #1/1/1900#
 Else
  MinDate = dDate
 End If
End Function


Public Function bitAnd(o1 As Long, o2 As Long) As Long
'Выполняет побитовое "И"
'#param o1: Первый операнд
'#param o2: Второй операнд
 bitAnd = o1 And o2
End Function
