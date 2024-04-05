Attribute VB_Name = "VBARegExp"
Option Compare Binary
 
Public Function GetRegExp(ByVal spPattern)
'Создает объект регулярного выражения.
'#param spPattern - Строка регулярного выражения. В дополнение к стандартному синтаксису в начало можно добавить следующие модификаторы:
' {*} \g - Глобальный поиск, в противном случае поиск будет остановлен на первом вхождении
' {*} \i - Игнорировать регистр при сравнении
' {*} \c - Сравнивать с учетом регистра (режим по умолчанию, можно не ставить)
' {*} \m - Мультистрочный режим. Отвечает за работу квантификаторов ^ и $.
'          Без модификатора соответствуют началу концу всей строки.
'          С модификатором соответствуют началу концу подстрок (разделенных по символу перевода строк).
  If IsNull(spPattern) Then spPattern = ""
  Static regexp, svPatternPrev
  If svPatternPrev = spPattern Then
    Set GetRegExp = regexp
  Else
    svPatternPrev = spPattern
    If IsEmpty(regexp) Then
      Set GetRegExp = CreateObject("VBScript.RegExp")
    Else
      Set GetRegExp = regexp
    End If
    GetRegExp.Global = False: GetRegExp.IgnoreCase = False: GetRegExp.Multiline = False
    Do While Left(spPattern, 1) = "\"
      Select Case LCase(Mid(spPattern, 2, 1))
        Case "g": GetRegExp.Global = True
        Case "i": GetRegExp.IgnoreCase = True
        Case "c": GetRegExp.IgnoreCase = False
        Case "m": GetRegExp.Multiline = True
        Case Else: Exit Do
      End Select
      spPattern = Mid(spPattern, 3)
    Loop
    GetRegExp.Pattern = spPattern
    Set regexp = GetRegExp
  End If
End Function
 
Public Function regexp_replace(spText, spPattern, sWhen) As Variant
'Заменяет строку по патерну на значение
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
'#param sWhen - Строка которой необходимо заменить значение
  If IsNull(spPattern) Or IsNull(spText) Or IsNull(sWhen) Then
    regexp_replace = Null
  Else
    regexp_replace = GetRegExp(spPattern).Replace(spText, sWhen)
  End If
End Function
 
 
Public Function regexp_substr(spText, spPattern) As Variant
'Извлекает из строки первое вхождение соответствующее регулярному выражени.
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
  If IsNull(spPattern) Or IsNull(spText) Then
    regexp_substr = Null
  Else
    Dim Match
    For Each Match In GetRegExp(spPattern).Execute(spText)
      If Match.SubMatches.Count = 0 Then
        regexp_substr = Match.Value
      Else
        regexp_substr = Match.SubMatches(0)
      End If
      Exit For
    Next
  End If
End Function
 
Public Function regexp_like(spText, spPattern) As Variant
'Проверяет строку на соответствие регулярному выражению
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
  If IsNull(spPattern) Or IsNull(spText) Then
    regexp_like = Null
  Else
    regexp_like = GetRegExp(spPattern).test(spText)
  End If
End Function
