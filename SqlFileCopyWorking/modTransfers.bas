Attribute VB_Name = "modDataTransfer"
Option Explicit
Declare Function timeGetTime Lib "winmm.dll" () As Long
Public sBuffer As String

' Declare RDO objects
Public Sourcetest As rdoConnection
Public SourceEr As rdoError
Public SourceEnv As rdoEnvironment
Public SourceCon As New rdoConnection
Public SourceQuery As New rdoQuery
Public SourceResult As rdoResultset
Public SourceRowBuf As Variant
Public SourceRowsReturned As Integer
Public TargetEr As rdoError
Public TargetEnv As rdoEnvironment
Public TargetCon As New rdoConnection
Public TargetQuery As New rdoQuery
Public TargetResult As rdoResultset
Public TargetRowBuf As Variant
Public TargetRowsReturned As Integer
Public MaintEr As rdoError
Public MaintEnv As rdoEnvironment
Public MaintCon As New rdoConnection
Public MaintQuery As New rdoQuery
Public MaintResult As rdoResultset
Public MaintRowBuf As Variant
Public MaintRowsReturned As Integer
'End of RDO declarations
Public iCount As Integer
Public iRecordCount As Long
Public iCurrentCount As Long
Public Progress
Public iRecordsAffected As Long

Global SourceConnectTest As Boolean
Global TargetConnectTest As Boolean
Public Function ConvertDate(sDate As Variant) As String

Dim sWorkDate

If CDate(sDate) Then
    sWorkDate = Mid(sDate, 1, 2) & "-" & GetMonth(Mid(sDate, 4, 2)) & "-" & Mid(sDate, 7, Len(sDate) - 6)
Else
    sWorkDate = "'null'"
End If

ConvertDate = sWorkDate
End Function
Public Function GetMonth(sMonth As Integer) As String
'array aMonths ("Jan","Feb","Mar","apr")

Dim aMonths As Variant
aMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
GetMonth = aMonths(sMonth - 1)

End Function


Public Function GetFields(sFilename As String, FormX As Form, ListX As ListBox, ConX As rdoConnection) As Boolean

'Dim RowBuf As Variant
Dim iFields As Integer
Dim ii As Integer
Dim sFieldName As String * 20
Dim sFieldType As String * 12

Set MaintQuery = Nothing
    With MaintQuery
        .Name = "GetRowsQuery"
        .SQL = "Select * from " & sFilename
        .RowsetSize = 1
        Set .ActiveConnection = SourceCon
        Set MaintResult = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With
    MaintResult.Requery
Dim Test
'Test = ConX.LastQueryResults.Status
DoEvents
'RowBuf = MaintResult.GetRows(2)      'Get the total number of records in the SourceFile

DoEvents

iFields = MaintResult.rdoColumns.Count
 
 For ii = 0 To iFields - 1
 sFieldName = MaintResult.rdoColumns(ii).Name
 sFieldType = GetType(MaintResult.rdoColumns(ii).Type)
    ListX.AddItem sFieldName & sFieldType & MaintResult.rdoColumns(ii).Size
Next ii

End Function
Public Function ClearFile(sFilename As String, FormX As Form, ConX As rdoConnection) As Boolean

On Error GoTo Clearfile_Error

If FormX.optClear(0).Value = True Then Exit Function

'Dim RowBuf As Variant
Dim iFields As Integer
Dim ii As Integer

Set MaintQuery = Nothing
    With MaintQuery
        .Name = "GetRowsQuery"
        If FormX.optClear(1).Value = True Then
            .SQL = "Truncate Table " & sFilename
        Else
            .SQL = "Delete from " & sFilename
        End If
        .RowsetSize = 1
        Set .ActiveConnection = ConX
        Set MaintResult = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With
    'MaintResult.Requery

DoEvents
FormX.lstResults.AddItem "Cleared " & sFilename
FormX.stb.Panels(2).Text = "Cleared " & sFilename
DoEvents
Exit Function
Clearfile_Error:
If FormX.optClear(1).Value = True Then
    MsgBox ("Failed to truncate " & sFilename)
Else
    MsgBox ("Failed to delete from " & sFilename)
End If

End Function
Public Function GetType(Index As Integer) As String
Select Case Index

Case 1
    GetType = "Char"
    
Case 2
    GetType = "Numeric"
    
Case 3
    GetType = "Decimal"
    
Case 4
    GetType = "Integer"
    
Case 5
    GetType = "Small Integer"
    
Case 6
    GetType = "Float"
    
Case 7
    GetType = "Real"
    
Case 8
    GetType = "Double"
    
Case 9
    GetType = "Date"
    
Case 10
    GetType = "Time"
    
Case 11
    GetType = "TimeStamp"
    
Case 12
    GetType = "VarChar"
    
Case -1
    GetType = "LongVarChar"

Case -2
    GetType = "Binary"

Case -3
    GetType = "VarBinary"

Case -4
    GetType = "LongVarBinary"

Case -5
    GetType = "BigInteger"
    
Case -6
    GetType = "TinyInteger"
    
Case -7
    GetType = "Bit"
    
End Select

End Function
Public Function Transfer(sSourceFile As String, sTargetFile As String, FormX As Form) As Boolean

'On Error GoTo transfer_Error

Dim TotalRecords
Dim Printline As String
Dim NewKey As String
Dim FileName
Dim iFields As Integer
Dim RowBuf As Variant
Dim RowsReturned As Long
Dim bUpdate As Boolean
Dim TargetSQL As String

With FormX
    DoEvents
    TotalRecords = 0
    Load pgb
    pgb.Show
    DoEvents
    .stb.Panels(2).Text = "Setting up the Data Environment"
    
    ' Get the Rowcount on the Source File
    Set SourceQuery = Nothing
    With SourceQuery
        .Name = "GetRowsQuery"
        .SQL = "Select count (*) from " & sSourceFile
        .RowsetSize = 1
        Set .ActiveConnection = SourceCon
        Set SourceResult = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With
    SourceResult.Requery
    DoEvents
    RowBuf = SourceResult.GetRows(5)      'Get the total number of records in the SourceFile
    iRecordCount = RowBuf(0, 0)
    .lstResults.AddItem iRecordCount & " selected from sourcefile for upload *"
    DoEvents
    'End of recordcount code
    
    'Start the Update
    iCurrentCount = 0
    'Set up the Source Environment
    Set SourceQuery = Nothing
    With SourceQuery
        .Name = "GetRowsQuery"
        .SQL = "Select * from " & sSourceFile
        .RowsetSize = 1
        Set .ActiveConnection = SourceCon
        Set SourceResult = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With

    .stb.Panels(2).Text = "Submitting the querry"

     'Now Execute the SQL statement and get the Records
     'We need to loop for every 50 lines

    SourceResult.Requery 'Get all the records from the Sourcefile
    DoEvents
    
    Dim i As Integer
    Do Until SourceResult.EOF
    .stb.Panels(2).Text = "Retreiving the Recordsets"
        SourceRowBuf = SourceResult.GetRows(1000)      'Get all the records from the source file
        RowsReturned = UBound(SourceRowBuf, 2) + 1

        For i = 0 To RowsReturned - 1
            iCurrentCount = iCurrentCount + 1
            Progress = pgb.Progress(iCurrentCount, iRecordCount)
            .stb.Panels(2).Text = "Processing for : " & SourceRowBuf(0, i)
            DoEvents
            
            TargetSQL = CreateSQL(SourceResult, SourceRowBuf, i, sTargetFile)
            bUpdate = DoUpdate(TargetSQL)

        Next i

    TotalRecords = TotalRecords + RowsReturned

    Loop
    'End of Update
    
    .stb.Panels(2).Text = "Upload Completed"
    DoEvents
    .lstResults.AddItem "Upload completed"
    Unload pgb
End With
Exit Function

transfer_Error:
MsgBox (Err.Description)
End Function
Public Function DoUpdate(sSQL As String) As Boolean
On Error GoTo Update_Error

Set TargetQuery = Nothing
    With TargetQuery
        .Name = "GetRowsQuery"
        .SQL = sSQL
        .RowsetSize = 1
        Set .ActiveConnection = TargetCon
        Set TargetResult = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With
    'TargetResult.Requery 'Get all the records from the Sourcefile
    
    DoEvents
    Exit Function
Update_Error:
    MsgBox (Err.Description)
End Function
Public Function CreateSQL(ResultSetX As rdoResultset, BuffX As Variant, iRowNumber As Integer, sTargetFile As String) As String
'This function will check the fields types , and then convert the data to the correct field formats
Dim sSQL As String
Dim iFields As Integer
Dim ii As Integer
Dim check

sSQL = "Insert into " & sTargetFile & " Values("


'Loop throug the fields to determine the field type
 iFields = ResultSetX.rdoColumns.Count
 For ii = 0 To iFields - 1
    Select Case ResultSetX.rdoColumns(ii).Type
    
    Case rdTypeCHAR, rdTypeVARCHAR, rdTypeLONGVARCHAR 'Strings
        If IsNull(BuffX(ii, iRowNumber)) Then
            sSQL = sSQL & "" & "Null" & ","
        Else
            sSQL = sSQL & "'" & BuffX(ii, iRowNumber) & "',"
        End If
    
    Case rdTypeDATE, rdTypeTIMESTAMP

        sSQL = sSQL & "'" & ConvertDate(BuffX(ii, iRowNumber)) & "',"
    
    Case Else
    'If IsNumeric(BuffX(ii, iRowNumber)) Then
        sSQL = sSQL & BuffX(ii, iRowNumber) & ","
    'Else
        'sSQL = sSQL & "'" & BuffX(ii, iRowNumber) & "',"
    'End If
    End Select
 
Next ii

sSQL = Left(sSQL, Len(sSQL) - 1)

sSQL = sSQL & ")"

CreateSQL = sSQL
End Function
Public Function SourceInitRDO(sConnectString, FormX As Form, ListX As ListBox) As Boolean
' This routine will initialize the RDO environment
SourceInitRDO = True
On Error GoTo mERROR
With FormX

    ' Now Connect to the RDO database
    .stb.Panels(2).Text = "Initializing...."
    .lstResults.AddItem "Initialising Source Environment"
    Set SourceEnv = rdoEnvironments(0)
    
    Set SourceCon = SourceEnv.OpenConnection(dsName:=sConnectString, _
         Prompt:=rdDriverCompleteRequired)
         
    .lstResults.AddItem "Attempting connection to " & sConnectString
         
    While SourceCon.StillConnecting
        .stb.Panels(2).Text = "Busy connecting to Source Database"
        DoEvents
    Wend
    
    .stb.Panels(2).Text = "Connected To Source"
    .lstResults.AddItem "Connected  to " & sConnectString

End With

Call ShowTest(FormX, ListX, SourceCon)

Exit Function

mERROR:
SourceInitRDO = False
End Function
Public Function TargetInitRDO(sConnectString, FormX As Form, ListX As ListBox) As Boolean
' This routine will initialize the RDO environment
TargetInitRDO = True
On Error GoTo mERROR
With FormX

    ' Now Connect to the RDO database
    .stb.Panels(2).Text = "Initializing...."
    .lstResults.AddItem "Initialising Target Environment"
    Set TargetEnv = rdoEnvironments(0)
    
    Set TargetCon = TargetEnv.OpenConnection(dsName:=sConnectString, _
         Prompt:=rdDriverCompleteRequired)
         
    .lstResults.AddItem "Attempting connection to " & sConnectString
         
    While TargetCon.StillConnecting
        .stb.Panels(2).Text = "Busy connecting to Target Database"
        DoEvents
    Wend
    
    .stb.Panels(2).Text = "Connected To Target"
    .lstResults.AddItem "Connected  to " & sConnectString

End With

Call ShowTest(FormX, ListX, TargetCon)

Exit Function

mERROR:
TargetInitRDO = False
End Function
Public Function ShowTest(FormX As Form, ListX As ListBox, rdoConnVar As rdoConnection)
Dim ii As Integer

For ii = 0 To rdoConnVar.rdoTables.Count - 1
    ListX.AddItem rdoConnVar.rdoTables(ii).Name
    DoEvents
Next ii
DoEvents

End Function

Public Function getResults(sSqlString As String) As Variant

Dim TotalRecords
Dim Printline As String
Dim NewKey As String
Dim FileName
Dim iFields As Integer

With frmMain
    DoEvents
    TotalRecords = 0
    Load pgb
    pgb.Show
    frmMain.Caption = "Submitting SQL to database"
    DoEvents
    .stb.Panels(2).Text = "Setting up the Data Environment"
    
    
    
    ' Set up the Environment to execute
    Set qy = Nothing
    With qy
        .Name = "GetRowsQuery"
        .SQL = sSqlString
        .RowsetSize = 1
        Set .ActiveConnection = cn
        Set rs = .OpenResultset(rdOpenKeyset, rdConcurRowVer)
    End With
    
    .stb.Panels(2).Text = "Submitting the querry"
    
    ' Now Execute the SQL statement and get the Records
    ' We need to loop for every 50 lines
    
    rs.Requery
    DoEvents
    
    'Do Until rs.EOF
    .stb.Panels(2).Text = "Retreiving the Recordsets"
        RowBuf = rs.GetRows(10)      'Get the next 6000 rows
        RowsReturned = UBound(RowBuf, 2) + 1
    
        For i = 0 To RowsReturned - 1
            Dummy = pgb.Progress(i, RowsReturned - 1)
            .stb.Panels(2).Text = "Processing for : " & RowBuf(0, i)
            DoEvents
            
            'Now , create the new record
            'NewKey = RowBuf(0, i)
            'NewKey = EnCrypt(NewKey, 10)
            'Printline = NewKey & "," & RowBuf(1, i) & "," & RowBuf(2, i)
            'Print #1, Printline ' Write the record to the text file
    
        Next i
    getResults = RowBuf
    TotalRecords = TotalRecords + RowsReturned
    
    iFields = rs.rdoColumns.Count

   For ii = 0 To iFields - 1
        MsgBox (rs.rdoColumns(ii).Name)
   Next ii
    
    'Loop
    .stb.Panels(2).Text = "Retreiving the Recordsets"
    .stb.Panels(2).Text = "Total Records Retreived : " & TotalRecords
    Unload pgb
End With
End Function

