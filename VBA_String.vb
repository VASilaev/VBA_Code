
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
