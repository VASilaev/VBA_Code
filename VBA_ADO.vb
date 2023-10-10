''Link: https://evona.nl/simple-parameterized-queries-using-ado-in-vba/

Public Function getConn() As ADODB.Connection
    'Returns an opened ADODB.Connection using a client-side cursor
    Set getConn = New ADODB.Connection
    getConn.CursorLocation = adUseClient
    getConn.Open adoString
End Function

Public Function getConn() As ADODB.Connection
    Set getConn = CurrentProject.Connection
End Function

Public Function GetParameterType(parameter As Variant) As ADODB.DataTypeEnum
    'Dynamically determine parameter type
    'Inspired by Gustav Brock's CSql
    Dim LongLong As Integer
    #If Win32 Then
        LongLong = 20
    #End If
    #If Win64 Then
        LongLong = VBA.vbLongLong
    #End If
    
    Select Case VarType(parameter)
        Case vbEmpty            '    0  Empty (uninitialized).
            GoTo UnsupportedType
        Case vbNull             '    1  Null (no valid data).
            GoTo UnsupportedType
        Case vbInteger          '    2  Integer.
            GetParameterType = adInteger
        Case vbLong             '    3  Long integer.
            GetParameterType = adInteger
        Case vbSingle           '    4  Single-precision floating-point number.
            GetParameterType = adSingle
        Case vbDouble           '    5  Double-precision floating-point number.
            GetParameterType = adDouble
        Case vbCurrency         '    6  Currency.
            GetParameterType = adDecimal
        Case vbDate             '    7  Date.
            GetParameterType = adDate
        Case vbString           '    8  String.
            GetParameterType = adVarWChar 'Assumes strings are not long
        Case vbObject           '    9  Object.
            GoTo UnsupportedType
        Case vbError            '   10  Error.
            GoTo UnsupportedType
        Case vbBoolean          '   11  Boolean.
            GetParameterType = adBoolean
        Case vbVariant          '   12  Variant (used only with arrays of variants).
            GoTo UnsupportedType
        Case vbDataObject       '   13  A data access object.
            GoTo UnsupportedType
        Case vbDecimal          '   14  Decimal.
            GetParameterType = adDecimal
        Case vbByte             '   17  Byte.
            GetParameterType = adChar
        Case LongLong           '   20  LongLong integer (Valid on 64-bit platforms only).
            GetParameterType = adBigInt
        Case vbUserDefinedType  '   36  Variants that contain user-defined types.
            GoTo UnsupportedType
        Case vbArray            '   8192  Array.
            GoTo UnsupportedType
        Case Else               '   For example array + variant = 8204
            GoTo UnsupportedType
    End Select
    Exit Function
UnsupportedType:
    'MsgBox "Unsupported input parameter going to SQL statement!"
    'Stop
    'Prod: text will likely work
    GetParameterType = adVarWChar
End Function

Public Function ExecuteParameters(sqlString As String, ParamArray Params() As Variant) As ADODB.Recordset
    'Executes a parameterized query using ADODB on the server using the input parameters, returns the recordset
    'Connection remains open, to close by ref counting when recordset expires
    Dim cmd As New ADODB.Command
    Dim inputParam As Variant
    cmd.ActiveConnection = getConn
    cmd.CommandText = sqlString
    For Each inputParam In Params
        Set inputParam = cmd.CreateParameter(, GetParameterType(inputParam), adParamInput, Len(Nz(inputParam, " ")), inputParam)
        cmd.Parameters.Append inputParam
    Next inputParam
    cmd.CommandType = adCmdText
    Set ExecuteParameters = cmd.Execute()
End Function

'Executes an action query:
ExecuteParameters "INSERT INTO Table1(Field1, Field2) VALUES (?, ?)", "David o'Connel says ""Hi!""", 15
'Opens up a recordset
Dim rs As ADODB.Recordset
Set rs = ExecuteParameters("SELECT * FROM Table1 WHERE ID = ? OR Name Like  '%' & ?", 15, "o'Connel")            
