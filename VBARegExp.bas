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
  If IsObject(spPattern) Then
    Set GetRegExp = spPattern
  Else
    Static regexp, svPatternPrev
    If Not IsEmpty(svPatternPrev) And svPatternPrev = spPattern Then
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
      On Error GoTo OnError
      GetRegExp.Pattern = spPattern
      Set regexp = GetRegExp
    End If
  End If
Exit Function
OnError:
  svPatternPrev = Empty
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
 
Public Function regexp_replace(spText, spPattern, sWhen) As Variant
'Заменяет строку по патерну на значение
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение с расширенным синтаксисом (см. {link #GetRegExp})
'#param sWhen - Строка которой необходимо заменить значение
  If IsNull(spPattern) Or IsNull(spText) Or IsNull(sWhen) Then
    regexp_replace = Null
  ElseIf IsObject(spPattern) Then
    regexp_replace = spPattern.Replace(spText, sWhen)
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
    Dim Match, re
    If IsObject(spPattern) Then Set re = spPattern Else Set re = GetRegExp(spPattern)
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
  ElseIf IsObject(spPattern) Then
    regexp_like = spPattern.test(spText)
  Else
    regexp_like = GetRegExp(spPattern).test(spText)
  End If
End Function


Public Function regexp_split(spText, spPattern) As Variant
'Разбивает текст на массив строк по разделителю заданному регулярным выражением
'#param spText - Проверяемая строка
'#param spPattern - Регулярное выражение разделителя с расширенным синтаксисом (см. {link #GetRegExp}). Для корректной работы в начало добавляется модификатор `\m`
  If IsNull(spPattern) Or IsNull(spText) Then
    regexp_split = Null
  Else
    Dim Match, Matches, re, aResult, nStart, i
    
    If IsObject(spPattern) Then Set re = spPattern Else Set re = GetRegExp("\g" & spPattern)
    Set Matches = re.Execute(spText)
    If Matches.Count = 0 Then
      regexp_split = Array(spText)
    Else
      aResult = Array()
      ReDim aResult(Matches.Count)
      nStart = 1
      For i = 0 To Matches.Count - 1
        aResult(i) = Mid(spText, nStart, Matches(i).FirstIndex - nStart + 1)
        nStart = Matches(i).FirstIndex + Matches(i).Length + 1
      Next
      aResult(i) = Mid(spText, nStart)
      regexp_split = aResult
    End If
  End If
End Function

Public Function regexp_filter(aspText, spPattern, Optional bpInclude = True) As Variant
'Возвращает массив, содержащий подмножество строкового массива на основе заданных условий фильтра.
'#param aspText - Одномерный массив искомых строк
'#param spPattern - Регулярное выражение условия фильтра (см. {link #GetRegExp}).
'#param bpInclude - Значение Boolean указывает, следует ли возвращать подстроки, включающие или исключающие по условию. Если `bpInclude` имеет значение `True`,
'`regexp_filter` возвращает подмножество массива, соответствующее регулярному выражению. Если `bpInclude` имеет значение False, `regexp_filter` возвращает подмножество массива, не отвечающее условию регулярного выражения.

  If IsNull(spPattern) Or IsNull(spText) Or IsNull(bpExclude) Then
    regexp_filter = Null
  ElseIf IsArray(aspText) Then
    Dim Match, Matches, re, aResult, nStart, i, j
    If IsObject(spPattern) Then Set re = spPattern Else Set re = GetRegExp(spPattern)
    i = 0
    aResult = Array()
    ReDim aResult(UBound(aspText))
    For j = 0 To UBound(aspText)
      If re.test(aspText(j)) Xor Not bpInclude Then
        aResult(i) = aspText(j)
        i = i + 1
      End If
    Next
    If i > 1 Then
      ReDim Preserve aResult(i - 1)
      regexp_filter = aResult
    ElseIf i = 1 Then
      regexp_filter = Array(aResult(0))
    Else
      regexp_filter = Array()
    End If
  ElseIf regexp_like(aspText, spPattern) Xor Not bpInclude Then
    regexp_filter = Array(aspText)
  Else
    regexp_filter = Array()
  End If
End Function
