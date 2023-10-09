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
