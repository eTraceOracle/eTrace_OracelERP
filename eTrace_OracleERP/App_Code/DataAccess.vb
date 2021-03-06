Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Xml
Imports System.Runtime.CompilerServices
Imports System.Data.SqlTypes
Imports System.Configuration
Imports System.Globalization

Public Class DataAccess
    Inherits ComponentModel.Component
    Implements System.IDisposable
    Private _SConn As SqlConnection
    Private _SComm As SqlCommand
    Private _Sda As SqlDataAdapter
    Private _SConnString As String = ConfigurationSettings.AppSettings.Item("eTraceDBConnString")
    Private _AMLConnString As String = ConfigurationSettings.AppSettings.Item("eTraceAMLConnString")
    Public _OConnString As String = System.Configuration.ConfigurationManager.AppSettings("SARTYI")  'SART3I  'SART4I SARTYI  TART1I 
    Private _OAscpString As String = System.Configuration.ConfigurationManager.AppSettings("cliff")  '("CLIFFS0")    
    Public _OConnString_EJIT As String = System.Configuration.ConfigurationManager.AppSettings("TART1I_AutoEJIT")  'PART1I
    Private _OConn As OracleConnection
    Private _OComm As OracleCommand
    Private _Oda As OracleDataAdapter

#Region "Connection"

    Public Sub New()
        '_OConn = New OracleConnection(_OConnString)                      'connect to Oracle
        '_SConn = New SqlConnection(_SConnString)                      'connect to SQL
    End Sub

#End Region

    '''''''' Methods  

#Region "Transaction"
    Public Function Trans(ByVal strSql As String()) As Boolean
        Dim sql As String
        Dim _STrans As SqlTransaction

        Try
            _SConn = New SqlConnection(_SConnString)
            _SConn.Open()
            _STrans = _SConn.BeginTransaction()
            Dim _SCmd As New SqlCommand
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Connection = _SConn
            _SCmd.Transaction = _STrans
            For Each sql In strSql
                _SCmd.CommandText = sql
                _SCmd.ExecuteNonQuery()
            Next
            _STrans.Commit()
            Return True
        Catch ex As Exception
            _STrans.Rollback()
            ErrorLogging("Call-Trans", "", sql & ", with error: " & ex.Message & ex.Source, "E")
            Return False
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try
    End Function

    Public Function TransSql(ByVal SqlTable As DataTable) As Boolean
        Dim sql As String
        Dim _STrans As SqlTransaction

        Try
            _SConn = New SqlConnection(_SConnString)
            _SConn.Open()
            _STrans = _SConn.BeginTransaction()
            Dim _SCmd As New SqlCommand
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Connection = _SConn
            _SCmd.Transaction = _STrans

            Dim i As Integer
            For i = 0 To SqlTable.Rows.Count - 1
                sql = SqlTable.Rows(i)(0).ToString
                _SCmd.CommandText = sql
                _SCmd.ExecuteNonQuery()
            Next
            _STrans.Commit()
            Return True
        Catch ex As Exception
            ErrorLogging("Call-TransSql", "", sql & ", with error: " & ex.Message & ex.Source, "E")
            _STrans.Rollback()
            Return False
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try
    End Function

    'Public Function BeginTrans() As Hashtable
    '    Dim htDBValue As New Hashtable
    '    Dim _STrans As SqlTransaction = _SConn.BeginTransaction
    '    htDBValue.Add("_SConn", _SConn)
    '    htDBValue.Add("_STrans", _STrans)
    '    Return htDBValue
    'End Function

    'Public Function EndTrans(ByVal htDBValue As Hashtable) As Object
    '    ' Dim EndTransR As Object=Object
    '    Dim _STrans As SqlTransaction = DirectCast(htDBValue.Item("_STrans"), SqlTransaction)
    '    Dim _SConn As SqlConnection = DirectCast(htDBValue.Item("_SConn"), SqlConnection)
    '    _STrans.Commit()
    '    _SConn.Close()
    '    Return EndTrans
    'End Function

    'Public Sub RollBack(ByVal htDBValue As Hashtable)
    '    Dim _STrans As SqlTransaction = DirectCast(htDBValue.Item("_STrans"), SqlTransaction)
    '    Dim _SConn As SqlConnection = DirectCast(htDBValue.Item("_SConn"), SqlConnection)
    '    _STrans.Rollback()
    '    _SConn.Close()
    'End Sub

#End Region


#Region "ExecuteDataSet"

    Public Function ExecuteDataSet(ByVal cmdText As String, ByVal ParamArray cmdParams As SqlParameter()) As DataSet
        Return Me.ExecuteDataSet(cmdText, CommandType.StoredProcedure, cmdParams)
    End Function

    Public Function ExecuteDataSet(ByVal cmdText As String, Optional ByVal cmdType As CommandType = 1) As DataSet
        Return Me.ExecuteDataSet(cmdText, cmdType, Nothing)
    End Function

    Public Function ExecuteDataSet(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray cmdParams As SqlParameter()) As DataSet
        Dim ds As New DataSet

        Try
            ds.Tables.Add(Me.ExecuteDataTable(cmdText, cmdType, cmdParams))
            Return ds
        Catch ex As Exception
            'ErrorLogging("ExecuteDataSet1", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function ExecuteDataSet(ByVal cmdText As String, ByVal tablename As String) As DataSet
        Dim strSql(1), keySql(1) As String
        strSql(0) = cmdText
        keySql(0) = tablename

        Try
            Return ExecuteDataSet(strSql, keySql)
        Catch ex As Exception
            ErrorLogging("ExecuteDataSet2", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        End Try

    End Function

    Public Function ExecuteDataset(ByVal strSql As String(), ByVal keySql As String()) As DataSet
        Dim a, b As Integer
        Dim cmdTxt As String = ""
        For b = 0 To strSql.Length - 1
            cmdTxt = cmdTxt & strSql(b) & ";"
        Next

        Try
            _SConn = New SqlConnection(_SConnString)
            Dim Da As SqlDataAdapter = New SqlDataAdapter(cmdTxt, _SConn)
            Da.TableMappings.Add(keySql(0), keySql(0))
            For a = 1 To keySql.Length - 1
                Da.TableMappings.Add(keySql(0) & a, keySql(a))
            Next
            Dim Ds As DataSet = New DataSet()
            Da.SelectCommand.CommandTimeout = TimeOut_M30
            Da.Fill(Ds, keySql(0))
            Return Ds
        Catch ex As Exception
            ErrorLogging("ExecuteDataset3", "", cmdTxt & ", with error: " & ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function ExecuteDataset(ByVal strSql As String, ByVal cmdType As Integer) As DataSet

        Try
            _SConn = New SqlConnection(_SConnString)
            Dim Da As SqlDataAdapter = New SqlDataAdapter(strSql, _SConn)
            Dim Ds As DataSet = New DataSet()
            Da.SelectCommand.CommandTimeout = TimeOut_M30
            Da.Fill(Ds)
            Return Ds
        Catch ex As Exception
            ErrorLogging("ExecuteDataset4", "", strSql & ", with error: " & ex.Message & ex.Source, "E")
        End Try
    End Function

#End Region

    Private Function ColumnEqual(ByVal A As Object, ByVal B As Object) As Boolean

        ' Compares two values to see if they are equal. Also compares DBNULL.Value. 
        ' Note: If your DataTable contains object fields, then you must extend this 
        ' function to handle them in a meaningful way if you intend to group on them. 

        If A Is DBNull.Value AndAlso B Is DBNull.Value Then
            ' both are DBNull.Value 
            Return True
        End If
        If A Is DBNull.Value OrElse B Is DBNull.Value Then
            ' only one is DBNull.Value 
            Return False
        End If
        ' value type standard comparison 
        Return (A.Equals(B))
    End Function

    Public Function SelectDistinct(ByVal TableName As String, ByVal SourceTable As DataTable, ByVal FieldName As String) As DataTable
        Dim dt As New DataTable(TableName)
        dt.Columns.Add(FieldName, SourceTable.Columns(FieldName).DataType)

        Try
            Dim LastValue As Object = Nothing
            For Each dr As DataRow In SourceTable.Select("", FieldName)
                If LastValue Is Nothing OrElse Not (ColumnEqual(LastValue, dr(FieldName))) Then
                    LastValue = dr(FieldName)
                    dt.Rows.Add(New Object() {LastValue})
                End If
            Next
            Return dt

        Catch ex As Exception
            ErrorLogging("SelectDistinct", "", ex.Message & ex.Source, "E")
        End Try
    End Function

#Region "CreateDataTable"
    Public Function CreateDataTable(ByVal TableName As String, ByVal ParamArray columnName As String()) As DataTable
        Dim dt As DataTable

        Try
            dt = New Data.DataTable(TableName)
            If (Not columnName Is Nothing) Then
                Dim column As String
                For Each column In columnName
                    dt.Columns.Add(New Data.DataColumn(column, System.Type.GetType("System.String")))
                Next
            End If
            Return dt
        Catch ex As Exception
            ErrorLogging("CreateDataTable", "", ex.Message & ex.Source, "E")
        End Try

    End Function
#End Region

#Region "ExecuteDataTable"

    Public Function ExecuteDataTable(ByVal cmdText As String, Optional ByVal cmdType As CommandType = 1) As DataTable
        Return Me.ExecuteDataTable(cmdText, cmdType, Nothing)
    End Function

    Public Function ExecuteDataTable(ByVal cmdText As String, ByVal ParamArray cmdParams As SqlParameter()) As DataTable
        Return Me.ExecuteDataTable(cmdText, CommandType.StoredProcedure, cmdParams)
    End Function

    Public Function ExecuteDataTable(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray cmdParams As SqlParameter()) As DataTable
        Try
            _SConn = New SqlConnection(_SConnString)
            Dim _SCmd As New SqlCommand(cmdText, _SConn)
            Dim _SApt As New SqlDataAdapter(_SCmd)
            Dim dt As New DataTable

            _SCmd.CommandType = cmdType
            _SCmd.CommandTimeout = TimeOut_M30

            If (Not cmdParams Is Nothing) Then
                Dim p As SqlParameter
                For Each p In cmdParams
                    _SCmd.Parameters.Add(p)
                Next
            End If
            _SApt.Fill(dt)
            _SConn.Close()
            Return dt

        Catch ex As Exception
            'ErrorLogging("ExecuteDataTable", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try

    End Function

    Public Function ExecuteDataTable(ByVal cmdText As String, ByVal DBConnType As Integer) As DataTable

        Try
            Dim DBStrString = ""
            If DBConnType = 0 Then                  'eTrace Database
                DBStrString = _SConnString
            Else                                    'Local eTraceAML Database
                DBStrString = _AMLConnString
            End If

            _SConn = New SqlConnection(DBStrString)
            Dim _SCmd As New SqlCommand(cmdText, _SConn)
            Dim _SApt As New SqlDataAdapter(_SCmd)
            Dim dt As New DataTable

            _SCmd.CommandType = 1
            _SCmd.CommandTimeout = TimeOut_M30

            _SApt.Fill(dt)
            _SConn.Close()
            Return dt

        Catch ex As Exception
            ErrorLogging("ExecuteDataTable", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try

    End Function
#End Region

#Region "UpdateDataSet"

    Public Shared Sub UpdateDataset(ByVal insertCommand As SqlCommand, ByVal deleteCommand As SqlCommand, ByVal updateCommand As SqlCommand, ByVal dataSet As DataSet, ByVal tableName As String)
        If insertCommand Is Nothing Then
            Throw New ArgumentNullException("insertCommand")
        End If
        If deleteCommand Is Nothing Then
            Throw New ArgumentNullException("deleteCommand")
        End If
        If updateCommand Is Nothing Then
            Throw New ArgumentNullException("updateCommand")
        End If
        If tableName Is Nothing OrElse tableName.Length = 0 Then
            Throw New ArgumentNullException("tableName")
        End If
        ' 创建SqlDataAdapter,当操作完成后释放.
        Using dataAdapter As New SqlDataAdapter()
            ' 设置数据适配器命令
            dataAdapter.UpdateCommand = updateCommand
            dataAdapter.InsertCommand = insertCommand
            dataAdapter.DeleteCommand = deleteCommand
            ' 更新数据集改变到数据库
            dataAdapter.Update(dataSet, tableName)
            ' 提交所有改变到数据集.
            dataSet.AcceptChanges()
        End Using
    End Sub

#End Region

#Region "ExecuteNonQuery"

    Public Function ExecuteNonQuery(ByVal cmdText As String, Optional ByVal cmdType As CommandType = 1) As Integer
        Return Me.ExecuteNonQuery(cmdText, cmdType, Nothing)
    End Function

    Public Function ExecuteNonQuery(ByVal cmdText As String, ByVal ParamArray cmdParams As SqlParameter()) As Integer
        Return Me.ExecuteNonQuery(cmdText, CommandType.StoredProcedure, cmdParams)
    End Function

    Public Function ExecuteNonQuery(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray cmdParams As SqlParameter()) As Integer
        _SConn = New SqlConnection(_SConnString)
        Dim _SCmd As New SqlCommand(cmdText, _SConn)
        Dim intRes As Integer

        Try
            _SConn.Open()
            _SCmd.CommandType = cmdType
            _SCmd.CommandTimeout = TimeOut_M60
            If (Not cmdParams Is Nothing) Then
                Dim p As SqlParameter
                For Each p In cmdParams
                    _SCmd.Parameters.Add(p)
                Next
            End If
            intRes = _SCmd.ExecuteNonQuery
            _SConn.Close()
            Return intRes

        Catch ex As Exception
            If Not cmdText.Contains("sp_ErrorLogging") Then
                ErrorLogging("ExecuteNonQuery", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
            End If
            Return 0
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try

    End Function
#End Region

#Region "ExecuteReader"

    Public Function ExecuteReader(ByVal cmdText As String, ByVal ParamArray cmdParams As SqlParameter()) As SqlDataReader
        Return Me.ExecuteReader(cmdText, CommandType.StoredProcedure, cmdParams)
    End Function

    Public Function ExecuteReader(ByVal cmdText As String, Optional ByVal cmdType As CommandType = 1) As SqlDataReader
        Return Me.ExecuteReader(cmdText, cmdType, Nothing)
    End Function

    Public Function ExecuteReader(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray cmdParams As SqlParameter()) As SqlDataReader
        _SConn = New SqlConnection(_SConnString)
        Dim _SCmd As New SqlCommand(cmdText, _SConn)
        _SCmd.CommandTimeout = TimeOut_M30
        Try
            _SConn.Open()
            _SCmd.CommandType = cmdType
            If (Not cmdParams Is Nothing) Then
                Dim p As SqlParameter
                For Each p In cmdParams
                    _SCmd.Parameters.Add(p)
                Next
            End If
            Return _SCmd.ExecuteReader()
        Catch ex As Exception
            ErrorLogging("ExecuteReader", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        Finally
            ''If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try
    End Function
#End Region

#Region "ExecuteScalar"
    Public Function ExecuteScalar(ByVal cmdText As String) As Object
        _SConn = New SqlConnection(_SConnString)
        Dim _SCmd As New SqlCommand(cmdText, _SConn)
        _SCmd.CommandTimeout = TimeOut_M30
        Dim obj As Object
        Dim strFlag As String = "0"
        Try
            _SConn.Open()
            strFlag = strFlag + "  1"
            obj = _SCmd.ExecuteScalar()
            strFlag = strFlag + "  2"
            _SConn.Close()
            strFlag = strFlag + "  3"
            Return obj
            strFlag = strFlag + "  4"
        Catch ex As Exception
            '' Ensure the cmdText not exceeded 1000 characters.
            '' Update by Jackson Huang
            If (cmdText.Length <= 1000) Then
                ErrorLogging("ExecuteScalar1", "", strFlag + "  " + cmdText & ", with error: " & ex.Message & ex.Source, "E")
            Else
                ErrorLogging("ExecuteScalar2", "", strFlag + "  " + cmdText.Substring(0, 1000) & ", with error: " & ex.Message & ex.Source, "E")
            End If
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try

    End Function

    Public Function ExecuteScalar(ByVal cmdText As String, ByVal ParamArray cmdParams As SqlParameter()) As Object
        Return Me.ExecuteScalar(cmdText, CommandType.StoredProcedure, cmdParams)
    End Function

    Public Function ExecuteScalar(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray cmdParams As SqlParameter()) As Object
        _SConn = New SqlConnection(_SConnString)
        Dim _SCmd As New SqlCommand(cmdText, _SConn)
        Try
            _SConn.Open()
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.CommandType = cmdType
            _SCmd.Parameters.Add("ReturnValue", SqlDbType.Variant).Direction = ParameterDirection.ReturnValue
            If (Not cmdParams Is Nothing) Then
                Dim p As SqlParameter
                For Each p In cmdParams
                    _SCmd.Parameters.Add(p)
                Next
            End If
            _SCmd.ExecuteNonQuery()
            Dim ReturnValue As Object = RuntimeHelpers.GetObjectValue(_SCmd.Parameters.Item("ReturnValue").Value)
            _SConn.Close()
            Return ReturnValue

        Catch ex As Exception
            ErrorLogging("ExecuteScalar2", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
        Finally
            If _SConn.State <> ConnectionState.Closed Then _SConn.Close()
        End Try
    End Function


    '' Add by Jackson Huang on 11/26/2013
    Public Function ExecuteScalar(ByVal cmdText As String, ByVal conn As SqlConnection) As Object

        If (conn Is Nothing) Then
            Throw New NullReferenceException("conn can't be null")
        End If

        ''Dim conn As SqlConnection = New SqlConnection(connStr)
        Dim cmd As New SqlCommand(cmdText, conn)
        cmd.CommandTimeout = TimeOut_M30
        Dim obj As Object
        Try
            conn.Open()
            obj = cmd.ExecuteScalar()
            conn.Close()
            Return obj
        Catch ex As Exception
            ErrorLoggingOTO("ExecuteScalar2", "", cmdText & ", with error: " & ex.Message & ex.Source, "E")
            ErrorLoggingOTO("ExecuteScalar2", "", ex.Message & ex.Source, "E")
        Finally
            If conn.State <> ConnectionState.Closed Then conn.Close()
        End Try

    End Function

#End Region


#Region "SQL tran"
    Public Function mySqlCommand() As SqlCommand
        _SComm = New SqlCommand()
        _SConn = New SqlConnection(_SConnString)
        _SComm.Connection = _SConn
        Return _SComm
    End Function
    Public Function Sda_Sele() As SqlDataAdapter
        _Sda = New SqlDataAdapter()
        _Sda.SelectCommand = mySqlCommand()
        Return _Sda
    End Function
    Public Function Sda_Insert() As SqlDataAdapter
        _Sda = New SqlDataAdapter()
        _Sda.InsertCommand = mySqlCommand()
        Return _Sda
    End Function
    Public Function Sda_UpdateComm() As SqlDataAdapter
        _Sda = New SqlDataAdapter()
        _Sda.UpdateCommand = mySqlCommand()
        Return _Sda
    End Function
    Public Function Sda_Update() As SqlDataAdapter
        _Sda = New SqlDataAdapter()
        _Sda.InsertCommand = mySqlCommand()
        _Sda.UpdateCommand = mySqlCommand()
        _Sda.DeleteCommand = mySqlCommand()
        Return _Sda
    End Function
#End Region


#Region "Oracle tran"
    'Add Oracle Command 
    Public Function OraCommand() As OracleCommand
        _OComm = New OracleCommand()
        _OConn = New OracleConnection(_OConnString)
        _OComm.Connection = _OConn
        Return _OComm
    End Function
    Public Function Oda_Sele() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.SelectCommand = OraCommand()
        Return _Oda
    End Function
    Public Function Oda_Insert() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = OraCommand()
        Return _Oda
    End Function
    Public Function Oda_UpdateComm() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.UpdateCommand = OraCommand()
        Return _Oda
    End Function
    Public Function Oda_Update() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = OraCommand()
        _Oda.UpdateCommand = OraCommand()
        _Oda.DeleteCommand = OraCommand()
        Return _Oda
    End Function

    ' Add Transaction Command  09/04/2009
    Public Function Ora_Command_Trans() As OracleCommand
        Dim connection As New OracleConnection(_OConnString)
        Dim command As OracleCommand = connection.CreateCommand()
        connection.Open()
        command.Transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted)
        Return command
    End Function
    Public Function Oda_Insert_Tran() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = Ora_Command_Trans()
        Return _Oda
    End Function
    Public Function Oda_UpdateComm_Tran() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.UpdateCommand = Ora_Command_Trans()
        Return _Oda
    End Function

    Public Function Oda_Update_Tran() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = Ora_Command_Trans()
        _Oda.UpdateCommand = Ora_Command_Trans()
        _Oda.DeleteCommand = Ora_Command_Trans()
        Return _Oda
    End Function

    'Add another Oracle Command for ASCP Connection only   11/11/2009
    Public Function AscpCommand() As OracleCommand
        _OComm = New OracleCommand()
        _OConn = New OracleConnection(_OAscpString)
        _OComm.Connection = _OConn
        Return _OComm
    End Function
    Public Function Ascp_Sele() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.SelectCommand = AscpCommand()
        Return _Oda
    End Function
#End Region

#Region "Auto EJIT"
    Public Function OraCommand_EJIT() As OracleCommand
        _OComm = New OracleCommand()
        _OConn = New OracleConnection(_OConnString_EJIT)
        _OComm.Connection = _OConn
        Return _OComm
    End Function
    Public Function Oda_Sele_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.SelectCommand = OraCommand_EJIT()
        Return _Oda
    End Function
    Public Function Oda_Insert_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = OraCommand_EJIT()
        Return _Oda
    End Function
    Public Function Oda_UpdateComm_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.UpdateCommand = OraCommand_EJIT()
        Return _Oda
    End Function
    Public Function Oda_Update_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = OraCommand_EJIT()
        _Oda.UpdateCommand = OraCommand_EJIT()
        _Oda.DeleteCommand = OraCommand_EJIT()
        Return _Oda
    End Function
    Public Function Ora_Command_Trans_EJIT() As OracleCommand
        Dim connection As New OracleConnection(_OConnString_EJIT)
        Dim command As OracleCommand = connection.CreateCommand()
        connection.Open()
        command.Transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted)
        Return command
    End Function
    Public Function Oda_Insert_Tran_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = Ora_Command_Trans_EJIT()
        Return _Oda
    End Function
    Public Function Oda_UpdateComm_Tran_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.UpdateCommand = Ora_Command_Trans_EJIT()
        Return _Oda
    End Function
    Public Function Oda_Update_Tran_EJIT() As OracleDataAdapter
        _Oda = New OracleDataAdapter()
        _Oda.InsertCommand = Ora_Command_Trans_EJIT()
        _Oda.UpdateCommand = Ora_Command_Trans_EJIT()
        _Oda.DeleteCommand = Ora_Command_Trans_EJIT()
        Return _Oda
    End Function
#End Region
    Public Sub Dispose() Implements System.IDisposable.Dispose
        GC.SuppressFinalize(True)
    End Sub

    Public Function ErrorLogging(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String, ByVal Category As String) As Boolean

        Dim ErrFlag As Integer
        Try
            'Filter special character "'" from Error Message to avoid SP execution error
            ErrMsg = ErrMsg.Replace("'", "''")

            Dim Sqlstr As String
            Sqlstr = String.Format("exec sp_ErrorLogging '{0}', N'{1}', N'{2}', N'{3}'", ModuleName, User, ErrMsg, Category)
            ErrFlag = ExecuteNonQuery(Sqlstr)

            If ErrFlag = -1 Then Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Error log for One to One database.
    ''' </summary>
    ''' <param name="ModuleName"></param>
    ''' <param name="User"></param>
    ''' <param name="ErrMsg"></param>
    ''' <returns></returns>
    ''' <remarks>Update by Jackson Huang on 12/02/2013</remarks>
    Public Function ErrorLoggingOTO(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String, ByVal Category As String) As Boolean

        '' Use stored procedure instead of SQL.
        '' Modify by Jackson Huang on 12/02/2013
        Dim ra As Integer
        Dim myConnect As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceOTOConnString"))
        ''Dim RecordErrorCommand As SqlClient.SqlCommand

        Try
            Using com = New SqlClient.SqlCommand("sp_ErrorLogging", myConnect)
                ' Uses store procedure.
                com.CommandType = CommandType.StoredProcedure

                ' Pass jobId to store procedure.
                com.Parameters.Add("@Module", SqlDbType.Char).Value = ModuleName
                com.Parameters.Add("@Category", SqlDbType.Char).Value = Category
                com.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = User
                com.Parameters.Add("@ErrorMsg", SqlDbType.NVarChar).Value = ErrMsg
                com.Connection.Open()
                ra = com.ExecuteNonQuery()
                com.Connection.Close()
            End Using
        Catch ex As Exception
            myConnect.Close()
            Return False
        End Try

        If (ra = -1) Then
            Return True
        End If

        Return False

    End Function

End Class

