
Function ReadBLOB(sSource As String, tRecordSet As Recordset, sField As String)
'Ñ÷èòûâàåò BLOB çíà÷åíèå èç ôàéëà â íàáîð äàííûõ è âîçâðàùåò ÷èñëî ñ÷èòàííûõ áàéò
'#param sSource: Èìÿ ôàéëà êóäà áóäåò çàïèñàíî çíà÷åíèå
'#param tRecordSet: Íàáîð äàííûõ ñïîçèöèîíèðîâàííûé íà íóæíîé ñòðîêå
'#param sField: Èìÿ ïîëÿ â êîòîðîå íóæíî ñîõðàíèòü äàííûå

    Dim NumBlocks As Integer, SourceFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim lngMeter As Long
    Dim FileData As String
    Dim byteData() As Byte
    Dim RetVal As Variant
    On Error GoTo Err_ReadBLOB
    SourceFile = FreeFile
    Open sSource For Binary Access Read As SourceFile
    FileLength = LOF(SourceFile)
    If FileLength = 0 Then
        ReadBLOB = 0
        Exit Function
    End If
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize
    lngMeter = FileLength \ 1000
    RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", lngMeter)
    If LeftOver > 0 Then
      ReDim byteData(0 To LeftOver - 1)
      Get SourceFile, , byteData
      tRecordSet(sField).AppendChunk (byteData)
    End If
    lngMeter = LeftOver \ 1000
    RetVal = SysCmd(acSysCmdUpdateMeter, lngMeter)
    ReDim byteData(0 To BlockSize - 1)
    For i = 1 To NumBlocks
        Get SourceFile, , byteData
        tRecordSet(sField).AppendChunk (byteData)
        lngMeter = BlockSize * i \ 1000
        RetVal = SysCmd(acSysCmdUpdateMeter, lngMeter)
    Next i
    RetVal = SysCmd(acSysCmdRemoveMeter)
    Close SourceFile
    ReadBLOB = FileLength
    Exit Function
Err_ReadBLOB:
    ReadBLOB = -Err
    Exit Function
End Function


Function WriteBLOB(tRecordSet As Recordset, sField As String, sDestination As String)
'Ñîõðàíÿåò BLOB çíà÷åíèå èç íàáîðà äàííûõ â ôàéë è âîçâðàùåò ÷èñëî çàïèñàííûõ áàéò
'#param tRecordSet: Íàáîð äàííûõ ñïîçèöèîíèðîâàííûé íà íóæíîé ñòðîêå
'#param sField: Èìÿ ïîëÿ èç êîòîðîãî íóæíî ñîõðàíèòü äàííûå
'#param sDestination: Èìÿ ôàéëà êóäà áóäåò çàïèñàíî çíà÷åíèå

    Dim NumBlocks As Integer, DestFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim lngMeter As Long
    Dim byteData() As Byte
    Dim RetVal As Variant
    On Error GoTo Err_WriteBLOB
    FileLength = tRecordSet(sField).FieldSize()
    If FileLength = 0 Then
        WriteBLOB = 0
        Exit Function
    End If
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize
    DestFile = FreeFile
    Open sDestination For Binary As DestFile
    lngMeter = FileLength \ 1000
    RetVal = SysCmd(acSysCmdInitMeter, "Writing BLOB", lngMeter)
    If LeftOver > 0 Then
      byteData() = tRecordSet(sField).GetChunk(0, LeftOver)
      Put DestFile, , byteData
    End If
    lngMeter = LeftOver \ 1000
    RetVal = SysCmd(acSysCmdUpdateMeter, lngMeter)
    For i = 1 To NumBlocks
        byteData() = tRecordSet(sField).GetChunk((i - 1) * BlockSize + LeftOver, BlockSize)
        Put DestFile, , byteData
        lngMeter = (i * BlockSize + LeftOver) \ 1000
        RetVal = SysCmd(acSysCmdUpdateMeter, lngMeter)
    Next i
    RetVal = SysCmd(acSysCmdRemoveMeter)
    Close DestFile
    WriteBLOB = FileLength
    Exit Function
Err_WriteBLOB:
    WriteBLOB = -Err
    Exit Function
End Function
      
