Option Compare Database

'SELECT
'  string_agg_get(
'    Max(
'      String_Agg(
'        Месяц,
'        Concat("таблица1.strAgg1", код Mod 2)
'      )
'    ),
'    "; "
'  ) As strAgg1
'FROM таблица1
'group by код mod 2




Dim nIdx As Integer
Dim tState As New Scripting.Dictionary
Dim tStateReverse As New Scripting.Dictionary
Dim tData As New Scripting.Dictionary

Sub test()
  Dim a
  a = QuickSort(Array(Array(Array("b"), "b Data"), Array(Array("d"), "d Data"), Array(Array("a"), "a Data"), Array(Array("ggg"), "ggg Data")))
End Sub

Public Function String_Agg(sData As Variant, sVarName As String, ParamArray aOrderBy()) As Integer
  'Debug.Print "sData = [" & sData & "] sVarName = [" & sVarName & "]"

  If IsNull(sData) Then
    String_Agg = 0
  Else
    String_Agg = tState(sVarName)
    If String_Agg = 0 Then
      nIdx = nIdx + 1
      String_Agg = nIdx
      tState(sVarName) = String_Agg
      
      If tData.Exists(String_Agg) Then
        If IsObject(tData(String_Agg)) Then tData(String_Agg).RemoveAll
        tData.Remove (String_Agg)
      End If
      
      If UBound(aOrderBy) = -1 Then
        'Без сортировки
        tStateReverse(String_Agg) = Array(sVarName, 0)
      Else
        tStateReverse(String_Agg) = Array(sVarName, 1)
      End If
      
      tData.Add String_Agg, New Scripting.Dictionary
      
    ElseIf Not IsObject(tData(String_Agg)) Then
      tData.Remove (String_Agg)
      tData.Add String_Agg, New Scripting.Dictionary
      
      If UBound(aOrderBy) = -1 Then
        'Без сортировки
        tStateReverse(String_Agg) = Array(sVarName, 0)
      Else
        tStateReverse(String_Agg) = Array(sVarName, 1)
      End If
    End If
    
    With tData(String_Agg)
      If UBound(aOrderBy) = -1 Then
        .Add .Count, sData
      Else
        Dim arr
        arr = aOrderBy
      
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = .Count
        .Add .Count, Array(arr, sData)
      End If
    End With
  End If
End Function


Public Function String_Agg_Get(Ind As Integer, sDelimeter As String, Optional isDistinct As Boolean) As Variant
  Dim aData, ItemData, j
  If Ind = 0 Then
    String_Agg_Get = ""
  ElseIf tData.Exists(Ind) Then
    If IsObject(tData(Ind)) Then
      If tStateReverse(Ind)(1) Then
        Dim itemKey, distinctDict As Scripting.Dictionary
        String_Agg_Get = Empty
        aData = QuickSort(tData(Ind).Items)
        
        Dim i As Integer
        
        For i = LBound(aData) To UBound(aData)
          ItemData = aData(i)(1)
          If isDistinct Then
            If distinctDict Is Nothing Then
              Set distinctDict = New Scripting.Dictionary
              distinctDict.Add ItemData, Empty
            Else
              If distinctDict.Exists(ItemData) Then
                GoTo dontAddItem
              Else
                distinctDict.Add ItemData, Empty
              End If
            End If
          End If
        
          If IsEmpty(String_Agg_Get) Then
            String_Agg_Get = ItemData
          Else
            String_Agg_Get = String_Agg_Get & sDelimeter & ItemData
          End If
dontAddItem:
        Next
        aData = Empty
      Else
        If isDistinct Then
          aData = tData(Ind).Items
          QuickSortInplace aData
          i = LBound(aData)
          Do While i < UBound(aData)
            j = i + 1
            Do While aData(j) = aData(i)
              aData(j) = "<REMOVE_ITEM>"
              j = j + 1
              If j > UBound(aData) Then Exit Do
            Loop
            i = j
          Loop
          aData = Filter(aData, "<REMOVE_ITEM>", False)
          String_Agg_Get = Join(aData, sDelimeter)
        Else
          String_Agg_Get = Join(tData(Ind).Items, sDelimeter)
        End If
        
      End If
    
      tData(Ind).RemoveAll
      tData(Ind) = String_Agg_Get
    Else
      String_Agg_Get = tData(Ind)
    End If
  Else
    String_Agg_Get = Null
  End If
End Function


Public Function ClearStringAggCache(Optional nIdx As Integer = -1)
  If nIdx = -1 Then
    tState.RemoveAll
    tStateReverse.RemoveAll
    Dim Item
    For Each Item In tData.Items
      If IsObject(Item) Then Item.RemoveAll
    Next
    tData.RemoveAll
  Else
    Dim sVarName As String
    sVarName = tStateReverse(nIdx)(0)
    tState.Remove (sVarName)
    tStateReverse.Remove (nIdx)
    If IsObject(tData(nIdx)) Then tData(nIdx).RemoveAll
    tData.Remove (nIdx)
  End If
End Function

Public Function QuickSort(vArray As Variant)
  QuickSort = vArray
  Call QuickSortInplace(QuickSort)
  Debug.Print "ok"
End Function

Private Function QuickSortCompare(ByRef a, ByRef b)
  If IsArray(a) Then
    If IsArray(b) Then
      If UBound(a) = -1 Then
        If UBound(b) = -1 Then
          Compare = 0
        Else
          Compare = -1
        End If
      ElseIf UBound(b) = -1 Then
        Compare = 1
      ElseIf IsArray(a(0)) Then
        If IsArray(b(0)) Then
          'Сравниваем два массива
          Dim i As Integer, HiA As Integer, HiB As Integer
          HiA = UBound(a(0))
          HiB = UBound(b(0))
          
          If HiA = HiB Then
            'Срвниваем массивы только одинаковой длины
            i = 0
            Do While i <= HiA And i <= HiB
              Compare = StrComp(a(0)(i), b(0)(i))
              If Compare <> 0 Then Exit Function
              i = i + 1
            Loop
          ElseIf HiA > HiB Then
            Compare = 1
          Else
            Compare = -1
          End If
        Else
          Compare = 1
        End If
      ElseIf IsArray(b(0)) Then
        'Массив всегда больше
        Compare = -1
      Else
        Compare = StrComp(a(0), b(0))
      End If
    ElseIf UBound(a) = -1 Then
      Compare = -1
    ElseIf IsArray(a(0)) Then
      Compare = StrComp(a(0)(0), b)
    Else
      Compare = StrComp(a(0), b)
    End If
  ElseIf IsArray(b) Then
    If UBound(b) = -1 Then
      Compare = 1
    ElseIf IsArray(b(0)) Then
      Compare = StrComp(a, b(0)(0))
    Else
      Compare = StrComp(a, b(0))
    End If
  Else
     Compare = StrComp(a, b)
  End If
End Function

Public Sub QuickSortInplace(ByRef vArray As Variant, Optional ByVal inLow As Long = -1, Optional ByVal inHi As Long = -1)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  If inLow = -1 Then inLow = LBound(vArray)
  If inHi = -1 Then inHi = UBound(vArray)
  tmpLow = inLow
  tmpHi = inHi
  pivot = vArray((inLow + inHi) \ 2)
  While (tmpLow <= tmpHi)
     While (QuickSortCompare(vArray(tmpLow), pivot) < 0 And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
     While (QuickSortCompare(pivot, vArray(tmpHi)) < 0 And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend
     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpSwap = Empty
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend
  If (inLow < tmpHi) Then QuickSortInplace vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSortInplace vArray, tmpLow, inHi
End Sub

Public Function Concat(ParamArray aStr() As Variant)
  On Error GoTo ManualCancat
  Concat = Join(aStr, "")
  Exit Function
ManualCancat:
  Concat = ""
  Dim Item
  For Each Item In aStr
    If IsNull(Item) Then Concat = Concat & "[NULL]" Else Concat = Concat & Item
  Next
End Function

