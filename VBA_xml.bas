Public Function XMLElement(spTag, spAttribute, spContent)
'Формирует один закрытый тэг
  XMLElement = "<" & spTag
  If spAttribute <> "" Then XMLElement = XMLElement & " " & spAttribute
  XMLElement = XMLElement & ">" & spContent & "</" & spTag & ">" & vbCrLf
End Function

Public Function XMLValue(sValue)
'Экранирует символы недопустимые в XML
  XMLValue = Replace(Replace(Replace(Replace(Replace(Replace(sValue, "&", "&amp;"), "«", """"), "»", """"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")
End Function
