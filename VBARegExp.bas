Attribute VB_Name = "VBARegExp"
Option Compare Binary

Public Function GetRegExp(spPattern)
'Создает объект регулярного выражения.
'#param spPattern - Строка регулярного выражения. В дополнение к стандартному синтаксису в начало можно добавить следующие модификаторы:
' {*} \g - Глобальный поиск, в противном случае поиск будет остановлен на первом вхождении
' {*} \i - Игнорировать регистр при сравнении
' {*} \c - Сравнивать с учетом регистра (режим по умолчанию, можно не ставить)
' {*} \m - Мультистрочный режим. Отвечает за работу квантификаторов ^ и $.
'          Без модификатора соответствуют началу концу всей строки.
'          С модификатором соответствуют началу концу подстрок (разделенных по символу перевода строк).

  Static regexp, svPatternPrev
  If svPatternPrev = spPattern Then
    Set GetRegExp = regexp
  Else
    Dim svPattern: svPattern = spPattern: Set GetRegExp = CreateObject("VBScript.RegExp")
    GetRegExp.Global = False: GetRegExp.IgnoreCase = False: GetRegExp.Multiline = False
    Do While Left(svPattern, 1) = "\"
      Select Case LCase(Mid(svPattern, 2, 1))
        Case "g": GetRegExp.Global = True
        Case "i": GetRegExp.IgnoreCase = True
        Case "c": GetRegExp.IgnoreCase = False
        Case "m": GetRegExp.Multiline = True
        Case Else: Exit Do
      End Select
      svPattern = Mid(svPattern, 3)
    Loop
    GetRegExp.Pattern = svPattern
    svPatternPrev = spPattern
    Set regexp = GetRegExp
  End If
End Function

Public Function regexp_replace(spText, spPattern, sWhen) As String
'Заменяет строку по патерну на значение
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
'#param sWhen - Строка которой необходимо заменить значение

  regexp_replace = GetRegExp(spPattern).Replace(spText, sWhen)
End Function


Public Function regexp_substr(spText, spPattern) As String
'Извлекает из строки первое вхождение соответствующее регулярному выражени.
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
  Dim Match
  For Each Match In GetRegExp(spPattern).Execute(spText)
    If Match.SubMatches.Count = 0 Then
      regexp_substr = Match.Value
    Else
      regexp_substr = Match.SubMatches(0)
    End If
    Exit For
  Next
End Function

Public Function regexp_like(spText, spPattern) As Boolean
'Проверяет строку на соответствие регулярному выражению
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
  regexp_like = GetRegExp(spPattern).test(spText)
End Function
