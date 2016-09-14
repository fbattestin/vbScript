# vbScript


' Before we can start you’ll need to add a reference to your VBA project:
' Microsoft ActiveX Data Objects x.x Library

Option Explicit
Private Conn As ADODB.Connection

Function ConnectToDB(Server As String, Database As String) As Boolean
 
    Set Conn = New ADODB.Connection
    On Error Resume Next
    
    Server = "GERCOR0603P\GESTAODB,1433"
    Database = "ADM_INFO"
    
    Conn.ConnectionString = "Provider=SQLOLEDB.1; Integrated Security=SSPI; Server=" & Server & "; Database=" & Database & ";"
    Conn.Open
    
    If Conn.State = 0 Then
        ConnectToDB = False
    Else
        ConnectToDB = True
    End If
 
End Function

Function Query(SQL As String)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field
 
    Dim Col As Long

    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        Col = 1
        For Each Field In recordSet.Fields
            Cells(1, Col) = Field.Name
            Col = Col + 1
        Next Field

        Cells(2, 1).CopyFromRecordset recordSet
        Set recordSet = Nothing
    End If
End Function

Public Sub Run()
 
    Dim SQL As String
    Dim Connected As Boolean
 
    SQL = "SELECT  SCHEMA_NAME(o.Schema_ID) + N'.' + o.NAME AS [Object Name], o.type_desc AS [Object Type]," & _
            "      i.name AS [Index Name], STATS_DATE(i.[object_id], i.index_id) AS [Statistics Date], " & _
            "      st.row_count, st.used_page_count,si.rowmodctr," & _
            "  CAST((CAST(si.rowmodctr AS DECIMAL(28,8))/CAST(st.row_count AS " & _
            " DECIMAL(28,2)) * 100.0)" & _
            " AS DECIMAL(28,2)) AS 'PCT_RowsChanged'" & _
            " FROM sys.objects AS o WITH (NOLOCK)" & _
            " INNER JOIN sys.indexes AS i WITH (NOLOCK)" & _
            " ON o.[object_id] = i.[object_id]" & _
            " INNER JOIN sys.sysindexes as si" & _
            " ON i.[object_id] = si.[id]" & _
            " AND i.index_id  = si.indid" & _
            " INNER JOIN sys.stats AS s WITH (NOLOCK)" & _
            " ON i.[object_id] = s.[object_id] " & _
            " AND i.index_id = s.stats_id" & _
            " INNER JOIN sys.dm_db_partition_stats AS st WITH (NOLOCK)" & _
            " ON o.[object_id] = st.[object_id]" & _
            " AND i.[index_id] = st.[index_id]" & _
            " WHERE o.[type] IN ('U', 'V')" & _
            " AND st.row_count > 0" & _
            " ORDER BY CAST((CAST(si.rowmodctr AS DECIMAL(28,8))/CAST(st.row_count AS" & _
            " DECIMAL(28,2)) * 100.0)" & _
            " AS DECIMAL(28,2)) DESC OPTION (RECOMPILE);  "

 
    Connected = ConnectToDB("SQL_SERVER", "DB_Name")
 
    If Connected Then
        Call Query(SQL)
        
        Conn.Close
    Else
        MsgBox "Huston we have a problem!"
    End If
 
End Sub
