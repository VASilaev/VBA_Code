Public Function getNumber(sMethodNumeration, ParamArray apSection())
  Dim aSection, Method, Row
  aSection = apSection
  
  Set Method = SelectOneRow("select id, trim(sFormat) as sFormat from tNumerationMethod where UCase(sName) = " & StrToSQL(UCase(sMethodNumeration)))
  
  If Method.Count = 0 Then
    Err.Raise 1000, , printf("Не удалось найти метод нумерации с кодом [{0}]", sMethodNumeration)
  Else
    DBEngine.Workspaces(0).BeginTrans
    On Error GoTo Rollback
    Set Row = SelectOneRow(printf("select id, nLastNumber from tNumeration as t where t.idNumerationMethod = {0} and t.sSection = {1}", Method("id"), StrToSQL(ArrayToLiteral(aSection))))
    If Row.Count = 0 Then
      getNumber = Replace(printf(Method("sFormat"), aSection), "{NUM}", 1)
      CurrentDb().Execute printf("insert into tNumeration(idNumerationMethod,nLastNumber,sSection) values ({0}, 1, {1})", Method("id"), StrToSQL(ArrayToLiteral(aSection)))
    Else
      getNumber = Replace(printf(Method("sFormat"), aSection), "{NUM}", Row("nLastNumber") + 1)
      CurrentDb().Execute printf("update tNumeration set nLastNumber = {1} where id = {0}", Row("id"), Row("nLastNumber") + 1)
    End If
    DBEngine.Workspaces(0).CommitTrans
  End If
  Exit Function
Rollback:
   On Error GoTo 0
   DBEngine.Workspaces(0).Rollback
End Function

'Создает необходимые таблицы
Public Sub InstallNumeratiomSystem()
  Dim vTbl, vFld, vDB
  
  Set vDB = CurrentDb()
  On Error GoTo onCreate
  Set vTbl = vDB.TableDefs("tNumeration")
  On Error GoTo 0
  Exit Sub
onCreate:
  On Error GoTo 0
  
  vDB.Execute "CREATE TABLE tNumerationMethod (id counter CONSTRAINT PK_NumerationMethod PRIMARY KEY, sName CHAR(255), sFormat CHAR(255));"
    
  vDB.Execute "CREATE TABLE tNumeration (id counter CONSTRAINT PK_Numeration PRIMARY KEY, idNumerationMethod long, nLastNumber long, sSection CHAR(255), CONSTRAINT FK_tNumeration_NumerationMethod FOREIGN KEY (idNumerationMethod) REFERENCES tNumerationMethod(id));"
    
  vDB.TableDefs.Refresh

End Sub
