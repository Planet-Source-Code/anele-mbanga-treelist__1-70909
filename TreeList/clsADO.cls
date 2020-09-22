VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public adoCn As ADODB.Connection
Private m_DatabaseName As String
Private m_CommandTimeOut As Long
Private m_ConnectionTimeout As Long
Private m_CursorLocation As Long
Private mvarKeepLog As Boolean
Private mvarLogFile As String
Private m_Table_FieldPrimaryKeyName As String
Private m_QueryString As String
Private intLog As Integer
Public Property Get QueryString() As String
    On Error Resume Next
    QueryString = m_QueryString
    Err.Clear
End Property
Public Property Let QueryString(ByVal Value As String)
    On Error Resume Next
    m_QueryString = Value
    Err.Clear
End Property
Public Property Get Table_FieldPrimaryKeyName() As String
    On Error Resume Next
    Table_FieldPrimaryKeyName = m_Table_FieldPrimaryKeyName
    Err.Clear
End Property
Public Property Get CursorLocation() As CursorLocationEnum
    On Error Resume Next
    CursorLocation = m_CursorLocation
    Err.Clear
End Property
Public Property Let CursorLocation(ByVal Value As CursorLocationEnum)
    On Error Resume Next
    m_CursorLocation = Value
    Err.Clear
End Property
Public Property Get ConnectionTimeout() As Long
    On Error Resume Next
    ConnectionTimeout = m_ConnectionTimeout
    Err.Clear
End Property
Public Property Let ConnectionTimeout(ByVal Value As Long)
    On Error Resume Next
    m_ConnectionTimeout = Value
    Err.Clear
End Property
Public Property Get CommandTimeOut() As Long
    On Error Resume Next
    CommandTimeOut = m_CommandTimeOut
    Err.Clear
End Property
Public Property Let CommandTimeOut(ByVal Value As Long)
    On Error Resume Next
    m_CommandTimeOut = Value
    Err.Clear
End Property
Public Property Get DatabaseName() As String
    On Error Resume Next
    DatabaseName = m_DatabaseName
    Err.Clear
End Property
Public Property Let DatabaseName(ByVal Value As String)
    On Error Resume Next
    m_DatabaseName = Value
    Err.Clear
End Property
Public Function OpenConnection(Optional ByVal DbName As String) As ADODB.ObjectStateEnum
    On Error Resume Next
    If Len(DbName) > 0 Then DatabaseName = DbName
    Set adoCn = New ADODB.Connection
    adoCn.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseName
    adoCn.CursorLocation = Me.CursorLocation
    adoCn.CommandTimeOut = Me.CommandTimeOut
    adoCn.ConnectionTimeout = Me.ConnectionTimeout
    adoCn.Mode = adModeReadWrite
    adoCn.Open
    OpenConnection = adoCn.State
    If Me.KeepLog = True Then
        intLog = FreeFile
        Open Me.LogFile For Append As #intLog
            Print #intLog, Me.DateTimeDone & ", Open Connection To Server " & DatabaseName
        End If
        Err.Clear
End Function
Public Function OpenRs(RsSource As String, Optional RsCursorType As CursorTypeEnum = adOpenStatic, Optional RsLockType As LockTypeEnum = adLockOptimistic, Optional RsCommandType As CommandTypeEnum = adCmdTable) As ADODB.Recordset
    On Error Resume Next
    Dim IsQuery As Boolean
    If Left$(LCase$(RsSource), 6) = "select" Then
        IsQuery = True
    Else
        RsSource = Replace$(RsSource, "[", "")
        RsSource = Replace$(RsSource, "]", "")
        RsSource = "[" & RsSource & "]"
        IsQuery = False
    End If
    Set OpenRs = New ADODB.Recordset
    OpenRs.CursorLocation = adoCn.CursorLocation
    OpenRs.Open RsSource, adoCn, RsCursorType, RsLockType, RsCommandType
    If IsQuery = True Then
        OpenRs.MoveLast
        OpenRs.MoveFirst
    End If
    Err.Clear
End Function
Public Function Table_NameFromSelect(ByVal strQuery As String) As String
    On Error Resume Next
    Dim fromPos As Long
    Dim lenStr As Long
    Dim lenCnt As Long
    Dim restStr As String
    Dim delimPos As Long
    Dim TbName As String
    fromPos = InStr(1, strQuery, " from ", vbTextCompare)
    If fromPos > 0 Then
        restStr = Trim$(Mid$(strQuery, fromPos + 5))
        ' check using the table Quote
        delimPos = InStr(1, restStr, "[", vbTextCompare)
        lenStr = Len(restStr)
        TbName = ""
        If delimPos > 0 Then
            For lenCnt = 2 To lenStr
                If Mid$(restStr, lenCnt, 1) = "]" Then
                    Exit For
                Else
                    TbName = TbName & Mid$(restStr, lenCnt, 1)
                End If
                Err.Clear
            Next
        Else
            For lenCnt = 1 To lenStr
                If Mid$(restStr, lenCnt, 1) = " " Then
                    Exit For
                Else
                    TbName = TbName & Mid$(restStr, lenCnt, 1)
                End If
                Err.Clear
            Next
        End If
        Table_NameFromSelect = Replace$(TbName, ";", "", , , vbTextCompare)
    Else
        Table_NameFromSelect = strQuery
    End If
    Err.Clear
End Function
Public Function RemNull(rsField As Variant) As String
    On Error Resume Next
    RemNull = rsField.Value & ""
    Err.Clear
End Function
Public Function OpenRsFieldNames(Rs As ADODB.Recordset, Optional QuoteEach As Boolean = False) As String
    On Error Resume Next
    Dim fldCnt As Long
    Dim fldTot As Long
    Dim fldHed As String
    Dim FldName As String
    fldHed = ""
    fldTot = Rs.Fields.Count - 1
    For fldCnt = 0 To fldTot
        FldName = Rs.Fields(fldCnt).Name
        If QuoteEach = True Then
            FldName = Chr$(34) & FldName & Chr$(34)
        End If
        If fldCnt = fldTot Then
            fldHed = fldHed & FldName
        Else
            fldHed = fldHed & FldName & ","
        End If
        Err.Clear
    Next
    OpenRsFieldNames = fldHed
    Err.Clear
End Function
Public Sub CloseConnection()
    On Error Resume Next
    adoCn.Close
    Err.Clear
End Sub
Public Property Get IsConnected() As Boolean
    On Error Resume Next
    Select Case adoCn.State
    Case 0
        IsConnected = False
    Case Else
        IsConnected = True
    End Select
    Err.Clear
End Property
Public Property Get State() As String
    On Error Resume Next
    Select Case adoCn.State
    Case adStateClosed
        State = "Closed"
    Case adStateOpen
        State = "Open"
    Case adStateExecuting
        State = "Executing"
    Case adStateFetching
        State = "Fetching"
    Case adStateConnecting
        State = "Connecting"
    End Select
    Err.Clear
End Property
Public Property Let KeepLog(ByVal vData As Boolean)
    On Error Resume Next
    ' if true a log file is kept by each execute statement
    mvarKeepLog = vData
    Err.Clear
End Property
Public Property Get KeepLog() As Boolean
    On Error Resume Next
    KeepLog = mvarKeepLog
    Err.Clear
End Property
Public Property Let LogFile(ByVal vData As String)
    On Error Resume Next
    ' establish the name of the log file
    mvarLogFile = vData
    Err.Clear
End Property
Public Property Get LogFile() As String
    On Error Resume Next
    LogFile = mvarLogFile
    Err.Clear
End Property
Private Sub Class_Initialize()
    On Error Resume Next
    Me.CommandTimeOut = 0
    Me.ConnectionTimeout = 15
    Me.CursorLocation = adUseClient
    Err.Clear
End Sub
Public Function Table_FieldNames(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the names of the fields of the table
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    strResults = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaColumns, Array(Empty, Empty, TbName))
    Do Until myRS.EOF
        strResults = strResults & myRS.Fields("column_name").Value & ","
        myRS.MoveNext
        Err.Clear
    Loop
    Table_FieldNames = RemDelim(strResults, ",")
    myRS.Close
    Err.Clear
End Function
Public Function Table_FieldIndexes(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the names of the indexes of the table
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    strResults = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaIndexes, Array(Empty, Empty, Empty, Empty, TbName))
    Do Until myRS.EOF
        strResults = strResults & myRS.Fields("column_name").Value & ","
        myRS.MoveNext
        Err.Clear
    Loop
    Table_FieldIndexes = MvRemoveDuplicates(strResults, ",")
    myRS.Close
    Err.Clear
End Function
Public Function Table_FieldSizes(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the sizes of the fields of the table
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    strResults = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaColumns, Array(Empty, Empty, TbName))
    Do Until myRS.EOF
        strResults = strResults & myRS.Fields("character_maximum_length").Value & ","
        myRS.MoveNext
        Err.Clear
    Loop
    Table_FieldSizes = RemDelim(strResults, ",")
    myRS.Close
    Err.Clear
End Function
Private Function RemDelim(ByVal Dataobj As String, ByVal Delimiter As String) As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delimiter)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delimiter
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function
Public Function MvRemoveDuplicates(ByVal StrMvString As String, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    ' returns a string from a string after removing all duplicated sub strings of a delimited string
    Dim spData() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim xCol As Collection
    Set xCol = New Collection
    spData = Split(StrMvString, Delim)
    spTot = UBound(spData)
    For spCnt = 0 To spTot
        spData(spCnt) = Trim$(spData(spCnt))
        If Len(spData(spCnt)) > 0 Then
            xCol.Add spData(spCnt), spData(spCnt)
        End If
        Err.Clear
    Next
    MvRemoveDuplicates = MvFromCollection(xCol, Delim)
    Err.Clear
End Function
Public Function MvFromCollection(objCollection As Collection, ByVal Delimiter As String) As String
    On Error Resume Next
    ' returns a delimited string based on a collection
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        If xCnt = xTot Then
            sRet = sRet & objCollection.Item(xCnt)
        Else
            sRet = sRet & objCollection.Item(xCnt) & Delimiter
        End If
        Err.Clear
    Next
    MvFromCollection = sRet
    Err.Clear
End Function
Public Function Table_FieldTypes(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the field types of the fields of the table
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    strResults = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaColumns, Array(Empty, Empty, TbName))
    Do Until myRS.EOF
        strResults = strResults & myRS.Fields("data_type").Value & ","
        myRS.MoveNext
        Err.Clear
    Loop
    Table_FieldTypes = RemDelim(strResults, ",")
    myRS.Close
    Err.Clear
End Function
Public Function Table_FieldPrimaryKey(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the names of the fields of the table
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    Dim strNames As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    strResults = ""
    strNames = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaPrimaryKeys, Array(Empty, Empty, TbName))
    Do Until myRS.EOF
        strResults = strResults & RemNull(myRS!column_name) & ","
        strNames = strNames & RemNull(myRS!pk_name) & ","
        myRS.MoveNext
        Err.Clear
    Loop
    strResults = MvRemoveDuplicates(strResults, ",")
    strNames = MvRemoveDuplicates(strNames, ",")
    m_Table_FieldPrimaryKeyName = strNames
    Table_FieldPrimaryKey = strResults
    myRS.Close
    Err.Clear
End Function
Public Function Table_FieldAutoIncrement(ByVal TbName As String) As String
    On Error Resume Next
    ' returns the names of the fields of the table that is autoincrement
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim strResults As String
    Dim strDefault As String
    Dim strColumnName As String
    If LCase$(Left$(TbName, 6)) = "select" Then
        TbName = Table_NameFromSelect(TbName)
    End If
    Table_FieldAutoIncrement = ""
    strResults = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaColumns, Array(Empty, Empty, TbName))
    Do Until myRS.EOF
        strDefault = LCase$(RemNull(myRS!column_default))
        strColumnName = RemNull(myRS!column_name)
        If strDefault = "('newsequentialid()')" Then
            strResults = strResults & strColumnName & ","
        End If
        myRS.MoveNext
        Err.Clear
    Loop
    myRS.Close
    Table_FieldAutoIncrement = RemDelim(strResults, ",")
    Err.Clear
End Function
Public Function Table_FieldIndexExists(ByVal TbName As String, ByVal strIdxName As String) As Boolean
    On Error Resume Next
    ' returns true / false if an index exists in a table
    Dim colPos As Long
    Dim colNames As String
    colNames = Table_FieldIndexes(TbName)
    colPos = MvSearch(colNames, strIdxName, ",")
    If colPos = 0 Then
        Table_FieldIndexExists = False
    Else
        Table_FieldIndexExists = True
    End If
    Err.Clear
End Function
Private Function MvSearch(ByVal StrMv As String, ByVal StrSearch As String, Delimiter As String) As Long
    On Error Resume Next
    Dim xValues() As String
    Dim xPos As Long
    xValues = Split(StrMv, Delimiter)
    xPos = ArraySearch(xValues, StrSearch)
    MvSearch = IIf((xPos = -1), 0, xPos + 1)
    Err.Clear
End Function
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    Dim arrayLow As Long
    ArrayTot = UBound(varArray)
    arrayLow = LBound(varArray)
    StrSearch = LCase$(Trim$(StrSearch))
    ArraySearch = -1
    For arrayCnt = arrayLow To ArrayTot
        strCur = LCase$(varArray(arrayCnt))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Public Function Table_FieldExists(ByVal TbName As String, ByVal strColumnName As String) As Boolean
    On Error Resume Next
    ' returns true or false if a field exists in a table
    Dim colPos As Long
    Dim colNames As String
    colNames = Table_FieldNames(TbName)
    colPos = MvSearch(colNames, strColumnName, ",")
    If colPos = 0 Then
        Table_FieldExists = False
    Else
        Table_FieldExists = True
    End If
    Err.Clear
End Function
Public Function Execute(ByVal mvarQuery As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    ' run an execute statement against the connection and
    ' log the entry of the log is being kept
    Set Execute = New ADODB.Recordset
    QueryString = mvarQuery
    If Me.KeepLog = True Then
        Print #intLog, Me.DateTimeDone & ", " & QueryString
    End If
    Set Execute = adoCn.Execute(mvarQuery)
    Err.Clear
    Exit Function
ErrorHandler:
    'Call Me.MyMsgBox
    Err.Clear
End Function
Public Sub Table_Create(ByVal dbTable As String, ByVal FldName As String, Optional ByVal fldType As String = "", Optional ByVal FldSize As String = "", Optional ByVal Fldidx As String = "", Optional ByVal FldAutoIncrement As String = "", Optional ByVal PrimaryFld As String = "")
    On Error Resume Next
    ' create a table
    Dim spFlds() As String
    Dim spType() As String
    Dim spSize() As String
    Dim spIndx() As String
    Dim spAuto() As String
    Dim totFld As Integer
    Dim totIdx As Integer
    Dim newCnt As Integer
    Dim tbQuery As String
    Dim typeStr As String
    Dim idxStr As String
    Fldidx = MvRemoveDuplicates(Fldidx, "," & PrimaryFld)
    FldAutoIncrement = MvRemoveDuplicates(FldAutoIncrement, ",")
    Call StrParse(spFlds, FldName, ",")
    Call StrParse(spType, fldType, ",")
    Call StrParse(spSize, FldSize, ",")
    Call StrParse(spIndx, Fldidx, ",")
    Call StrParse(spAuto, FldAutoIncrement, ",")
    ArrayTrimItems spFlds
    ArrayTrimItems spType
    ArrayTrimItems spSize
    ArrayTrimItems spIndx
    ArrayTrimItems spAuto
    totFld = UBound(spFlds)
    totIdx = UBound(spIndx)
    ReDim Preserve spType(totFld)
    ReDim Preserve spSize(totFld)
    dbTable = Iconv(dbTable, "t")
    tbQuery = "CREATE TABLE [" & dbTable & "] ("
    idxStr = ""
    For newCnt = 1 To totFld
        spType(newCnt) = Trim$(spType(newCnt))
        spFlds(newCnt) = Trim$(spFlds(newCnt))
        spSize(newCnt) = Trim$(spSize(newCnt))
        If Len(spType(newCnt)) = 0 Then
            spType(newCnt) = "Text"
        End If
        If Len(spSize(newCnt)) = 0 Then
            spSize(newCnt) = "255"
        End If
        Select Case LCase$(spType(newCnt))
        Case "image", "picture", "general": typeStr = "varbinary(max)"
        Case "nvarchar": typeStr = "nvarchar(" & spSize(newCnt) & ")"
        Case "boolean", "logical":          typeStr = "bit"
        Case "varbinary": typeStr = "varbinary(" & spSize(newCnt) & ")"
        Case "char": typeStr = "char(" & spSize(newCnt) & ")"
        Case "varchar": typeStr = "varchar(" & spSize(newCnt) & ")"
        Case "currency", "smallmoney":      typeStr = "money"
        Case "date":        typeStr = "datetime"
        Case "time":        typeStr = "datetime"
        Case "timestamp": typeStr = "datetime"
        Case "smalldatetime":        typeStr = "datetime"
        Case "double":         typeStr = "float"
        Case "real":        typeStr = "real"
        Case "integer", "int":        typeStr = "smallint"
        Case "long": typeStr = "int"
        Case "memo", "longtext", "long text":         typeStr = "nvarchar(max)"
        Case "single":             typeStr = "tinyint"
        Case "text":          typeStr = "varchar(" & spSize(newCnt) & ")"
        Case "ntext": typeStr = "nvarchar(max)"
        Case "xml": typeStr = "nvarchar(max)"
        Case ""
            typeStr = "varchar(255)"
        Case Else
            typeStr = LCase$(spType(newCnt))
        End Select
        If MvSearch(FldAutoIncrement, spFlds(newCnt), ",") > 0 Then
            typeStr = "uniqueidentifier DEFAULT NEWSEQUENTIALID()"
        End If
        If MvSearch(PrimaryFld, spFlds(newCnt), ",") > 0 Then
            typeStr = typeStr & " NOT NULL"
        End If
        tbQuery = tbQuery & "[" & spFlds(newCnt) & "] " & typeStr & ","
        Err.Clear
    Next
    tbQuery = RemDelim(tbQuery, ",") & ")"
    Me.Table_Delete dbTable
    Call Me.Execute(tbQuery)
    For newCnt = 1 To totIdx
        Fldidx = "create index [" & spIndx(newCnt) & "] on [" & dbTable & "] (" & spIndx(newCnt) & ");"
        Call Me.Execute(Fldidx)
        Err.Clear
    Next
    ' create the primary key
    If Len(PrimaryFld) > 0 Then
        PrimaryFld = MvField(PrimaryFld, 1, ",")
        QueryString = "ALTER TABLE [" & dbTable & "] ADD CONSTRAINT [PK_" & dbTable & "_" & PrimaryFld & "] PRIMARY KEY (" & PrimaryFld & ");"
        Call Me.Execute(QueryString)
    End If
    Err.Clear
End Sub
Public Function Table_Delete(ByVal TbName As String) As Boolean
    On Error Resume Next
    ' delete specified table
    If Len(TbName) = 0 Then
        Table_Delete = False
    Else
        If Table_Exists(TbName) = True Then Me.Execute "DROP TABLE [" & TbName & "];"
        If Table_Exists(TbName) = True Then
            Table_Delete = False
        Else
            Table_Delete = True
        End If
    End If
    Err.Clear
End Function
Public Sub Tables_Delete(ParamArray TablesToDelete())
    On Error Resume Next
    ' delete a collection of tables
    Dim Item As Variant
    For Each Item In TablesToDelete
        Call Table_Delete(CStr(Item))
        DoEvents
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function Table_Exists(ByVal strTbName As String) As Boolean
    On Error Resume Next
    ' returns existence of a table
    Dim strCurrent As String
    Dim dPos As Long
    strCurrent = Me.Table_Names
    dPos = MvSearch(strCurrent, strTbName, ",")
    If dPos = 0 Then
        Table_Exists = False
    Else
        Table_Exists = True
    End If
    Err.Clear
End Function
Public Property Get DateTimeDone() As String
    On Error Resume Next
    ' exact time now
    DateTimeDone = Format$(Now, "dd/mm/yyyy hh:mm:ss ampm")
    Err.Clear
End Property
Public Function ComputerName() As String
    On Error Resume Next
    ' return computer name
    ComputerName = VBA.Environ$("COMPUTERNAME")
    Err.Clear
End Function
Public Function Table_Names(Optional ByVal bShowSchema As Boolean = False, Optional ByVal bShowSystem As Boolean = False) As String
    On Error Resume Next
    ' returns all the names of the table types in the current database
    Dim myRS As ADODB.Recordset
    Set myRS = New ADODB.Recordset
    Dim mStr As String
    Dim mSource As String
    Dim mType As String
    Dim mName As String
    Dim tbType As String
    mStr = ""
    tbType = ""
    Set myRS = adoCn.OpenSchema(ADODB.adSchemaTables)
    Do Until myRS.EOF
        mSource = LCase$(RemNull(myRS!table_schema))
        mType = LCase$(RemNull(myRS!table_type))
        mName = RemNull(myRS!table_name)
        Select Case LCase$(mType)
        Case "system table"
            If bShowSystem = True Then
                tbType = tbType & mName & ","
            End If
        Case "access table"
            If bShowSchema = True Then
                tbType = tbType & mName & ","
            End If
        Case "table", "view"
            tbType = tbType & mName & ","
        End Select
        myRS.MoveNext
        Err.Clear
    Loop
    Table_Names = RemDelim(tbType, ",")
    myRS.Close
    Err.Clear
End Function
Public Function Table_FieldCreateIndex(ByVal TbName As String, ByVal IdxName As String) As Boolean
    On Error Resume Next
    Dim fExists As Boolean
    fExists = Table_FieldIndexExists(TbName, IdxName)
    If fExists = False Then
        ' create an index in a table
        QueryString = "create index [" & IdxName & "] on [" & TbName & "] (" & IdxName & ");"
        Call Execute(QueryString)
        DoEvents
    End If
    Table_FieldCreateIndex = Table_FieldIndexExists(TbName, IdxName)
    Err.Clear
End Function
Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delimiter As String) As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    varArray = Split(strText, Delimiter)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
Public Sub ArrayTrimItems(varArray() As String)
    On Error Resume Next
    ' trim all array elements
    Dim uArray As Long
    Dim cArray As Long
    Dim lArray As Long
    uArray = UBound(varArray)
    lArray = LBound(varArray)
    For cArray = lArray To uArray
        varArray(cArray) = Trim$(varArray(cArray))
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function Iconv(ByVal sValue As String, Optional ByVal sFormat As String = "") As String
    On Error Resume Next
    ' remove characters specified below from a string
    Dim sRslt As String
    Dim I As Long
    Dim Ch As String
    Dim L As Long
    Dim sN As String
    sRslt = sValue
    Select Case UCase$(sFormat)
    Case ""
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "(", "")
        sRslt = Replace$(sRslt, ")", "")
        sRslt = Replace$(sRslt, "~", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "@", "")
        sRslt = Replace$(sRslt, "#", "")
        sRslt = Replace$(sRslt, "$", "")
        sRslt = Replace$(sRslt, "%", "")
        sRslt = Replace$(sRslt, "^", "")
        sRslt = Replace$(sRslt, "&", "")
        sRslt = Replace$(sRslt, "*", "")
        sRslt = Replace$(sRslt, "_", "")
        sRslt = Replace$(sRslt, "-", "")
        sRslt = Replace$(sRslt, "=", "")
        sRslt = Replace$(sRslt, "|", "")
        sRslt = Replace$(sRslt, "\", "")
        sRslt = Replace$(sRslt, ":", "")
        sRslt = Replace$(sRslt, ";", "")
        sRslt = Replace$(sRslt, "<", "")
        sRslt = Replace$(sRslt, ">", "")
        sRslt = Replace$(sRslt, "?", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "+", "")
        sRslt = Replace$(sRslt, "{", "")
        sRslt = Replace$(sRslt, "}", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, Chr$(34), "")
    Case "Q"
        sRslt = Replace$(sRslt, "''", "")
        sRslt = Replace$(sRslt, "'", "")
    Case "F"
        sRslt = Replace$(sRslt, "/", "%")
        sRslt = Replace$(sRslt, "\", "%")
        sRslt = Replace$(sRslt, "|", "%")
    Case "C"
        sRslt = Replace$(sRslt, ",", "")
    Case "M"
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, ".", "")
    Case "S"
        L = Len(sRslt)
        sRslt = sRslt
        If L = 0 Then
            Err.Clear
            Exit Function
        End If
        sN = ""
        For I = 1 To L
            Ch = Mid$(sRslt, I, 1)
            If Ch = " " Then
                sN = sN & Ch
            End If
            If Ch >= "a" Then
                If Ch <= "z" Then
                    sN = sN & Ch
                End If
            End If
            If Ch >= "A" Then
                If Ch <= "Z" Then
                    sN = sN & Ch
                End If
            End If
            Err.Clear
        Next
        sRslt = sN
    Case "T"
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, Chr$(34), "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, ",", "")
    End Select
    Iconv = sRslt
    Err.Clear
End Function
Public Function MvField(ByVal strData As String, Optional ByVal fldPos As Long = 1, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    ' returns a substring from a delimted string
    Dim spData() As String
    Dim spCnt As Long
    MvField = ""
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function
