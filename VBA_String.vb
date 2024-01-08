
Public Function Translate(Text As String) As String
'Транслитерация кириллицы
'#param Text: Исходный текст

 Static rus, eng
 Dim i, j
 If IsEmpty(eng) Then
   rus = "аАбБвВгГдДеЕёЁжЖзЗиИйЙкКлЛмМнНоОпПрРсСтТуУфФхХцЦчЧшШщЩъЪыЫьЬэЭюЮяЯ"
   eng = Array("a", "A", "b", "B", "v", "V", "g", "G", "d", "D", "e", "E", "yo", "Yo", "zh", "Zh", "z", "Z", "i", "I", "j", "J", "k", "K", "l", "L", "m", "M", "n", "N", "o", "O", "p", "P", "r", "R", "s", "S", "t", "T", "u", "U", "f", "F", "kh", "Kh", "ts", "Ts", "ch", "Ch", "sh", "Sh", "sch", "Sch", "", "", "y", "Y", "", "", "e", "E", "yu", "Yu", "ya", "Ya")
 End If
 
 For i = 1 To Len(Text)
  j = InStr(1, rus, Mid(Text, i, 1), vbBinaryCompare)
  If j > 0 Then
   Translate = Translate & eng(j - 1)
  Else
   Translate = Translate & Mid(Text, i, 1)
  End If
 Next
End Function

Function decode(spExpression As Variant, ParamArray apArgs() As Variant) As Variant
'Первым DECODE принимает параметр, значение которого будет сравниваться по очереди со списком значений,
'и в случае, когда он совпадет с одним из перечисленных, будет возвращен соответствующий результат.
'Если совпадений не найдено, будет возвращено значение по-умолчанию. Если значение по-умолчанию не указано, будет возвращен NULL.
'`DECODE( expression , search , result [, search , result]... [, default] )`
'#param expression: выражение для сравнения
'#param search: значение, которое сравнивается с expression. В данной реализации Null = Null возвращает так же Истину
'#param result: значение, возвращаемое, если выражение expression совпало с искомым search
'#param default: значение по-умолчанию (необязательный)
 Dim bvIsNull As Boolean, i
 bvIsNull = IsNull(spExpression)
 decode = Null
 For i = 0 To UBound(apArgs) Step 2
  If i = UBound(apArgs) Then
   decode = apArgs(i)
   Exit Function
  ElseIf apArgs(i) = spExpression Or (bvIsNull And IsNull(apArgs(i))) Then
   decode = apArgs(i + 1)
   Exit Function
  End If
 Next
End Function
