Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Data.OracleClient
Imports System.Web.Services
Imports System.Web.Configuration
Imports System.IO
Imports System.Net.Mail
Imports System.IO.Compression

#Region "PublicData"

Public Structure ERPLogin
    Public Server As String
    Public User As String
    Public CHNName As String
    Public PWD As String
    Public OrgCode As String
    Public OrgID As String
    Public Application As String                  'PO / Inv / WIP / HH
    Public UserID As String
    Public UserType As String
    Public UserDept As String
    Public ResetFlag As Boolean
    Public Printer As String
    Public ErrorMsg As String
    Public AppID_PO As String
    Public RespID_PO As String
    Public AppID_Inv As String
    Public RespID_Inv As String
    Public AppID_WIP As String
    Public RespID_WIP As String
    Public AppID_KB As String
    Public RespID_KB As String
    Public ClientVersion As String
    Public ProductionLine As String
End Structure

Public Structure UserData
    Public User As String
    Public FirstName As String
    Public LastName As String
    Public OrgID As String
    Public UserID As String
    Public UserType As String
    Public UserDept As String
    Public ResetFlag As Boolean
    Public AppID_PO As String
    Public RespID_PO As String
    Public AppID_Inv As String
    Public RespID_Inv As String
    Public AppID_WIP As String
    Public RespID_WIP As String
    Public AppID_KB As String
    Public RespID_KB As String
    Public Printer As String
    Public Server As String
    Public PropVersion As String        ' Version of the process properties
    Public MinClientVersion As String
    Public RecClientVersion As String
    Public TransactionID As DataSet     ' Added by Andyking for User Authorization Control
    Public ErrorMsg As String
    Public ProductionLine As String
End Structure

Public Structure AccessCard
    Public AccessCardID As String
    Public EmplogeeID As String
    Public CHNName As String
    Public Name As String
    Public Dept As String
End Structure

Public Structure POLabel
    Public CLID As String
    Public OrgCode As String
    Public PCBA As String
    Public POID As String
    Public Qty As String
    Public UOM As String
    Public Rev As String
    Public RecDate As String
    Public SONo As String
    Public SOLine As String
    Public LotNo As String
    Public DateCode As String
    Public MPN As String
    Public RoHS As String
End Structure

Public Structure ItemRevList
    Dim Flag As String
    Dim Msg As String
    Dim RevList As String
End Structure

Public Structure ItemType
    Dim Type As String
    Dim Flag As String
    Dim Msg As String
End Structure

Public Structure CLIDSlot
    Dim Slot As String
    Dim OrgCode As String
    Dim SubInv As String
    Dim Locator As String
    Dim SlotType As String
    Dim SlotCLID As String
    Dim SlotMsg As String
End Structure

Public Structure CH09Label
    Public CLID As String
    Public CHCode As String
    Public Qty As String
    Public VendorID As String
    Public DateCode As String
    Public LotNo As String
    Public Rev As String
    Public MSL As String
End Structure

#End Region



''' <summary>
''' Base Response Class of webService
''' </summary>
Public Class ResponseBase
    ''' <summary>
    ''' indicate if webfunction run succeed
    ''' </summary>
    Public IsSucess As Boolean
    ''' <summary>
    ''' decripton of e
    ''' </summary>
    Public ErrorMessage As String

End Class
''' <summary>
''' response class of CopyProductInfo
''' </summary>

<System.Xml.Serialization.XmlRoot()>
Public Class CloneProductInfo
    Inherits ResponseBase
End Class
Public Class PublicFunction
    Private CLIDLabelFile As String = "D:\eTrace\CLMatLabel.lab"  'For zebra Printer QL420
    Private POLabelFile As String = "D:\eTrace\POLabel.lab"
    Private ESDImage As String = "D:\eTrace\ESD.jpg"
    Private BlankESD As String = "D:\eTrace\BlankESD.jpg"
    Private CH09LabelFile As String = "D:\eTrace\CH09Label.lab"


    Public Function GetDataAccess() As DataAccess
        Return (New DataAccess())
    End Function

    '#Region "Connection"
    '    Private _ConnString As String
    '    Private _OConn As OracleConnection
    '    Private _OComm As OracleCommand
    '    Private _Oda As OracleDataAdapter

    '    Public Sub New()
    '        ' _ConnString = System.Configuration.ConfigurationManager.AppSettings("zelda10")
    '        _ConnString = System.Configuration.ConfigurationManager.AppSettings("CAROLD3")
    '        _OConn = New OracleConnection(_ConnString)
    '    End Sub
    '    Public Function OraCommand() As OracleCommand
    '        _OComm = New OracleCommand()
    '        _OComm.Connection = _OConn
    '        Return _OComm
    '    End Function
    '    Public Function Oda_Sele() As OracleDataAdapter
    '        _Oda = New OracleDataAdapter()
    '        _Oda.SelectCommand = OraCommand()
    '        Return _Oda
    '    End Function
    '    Public Function Oda_Insert() As OracleDataAdapter
    '        _Oda = New OracleDataAdapter()
    '        _Oda.InsertCommand = OraCommand()
    '        Return _Oda
    '    End Function
    '    Public Function Oda_Update() As OracleDataAdapter
    '        _Oda = New OracleDataAdapter()
    '        _Oda.InsertCommand = OraCommand()
    '        _Oda.UpdateCommand = OraCommand()
    '        _Oda.DeleteCommand = OraCommand()
    '        Return _Oda
    '    End Function

    '    ' Add Transaction Command  09/04/2009
    '    Public Function Ora_Trans_Command() As OracleCommand
    '        Dim connection As New OracleConnection(_connString)
    '        Dim command As OracleCommand = connection.CreateCommand()
    '        command.Transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted)
    '        Return command
    '    End Function
    '    Public Function ODA_Insert_Tran() As OracleDataAdapter
    '        _oda = New OracleDataAdapter()
    '        _oda.InsertCommand = Ora_Trans_Command()
    '        Return _oda
    '    End Function

    '    Public Function ODA_Update_Tran() As OracleDataAdapter
    '        _oda = New OracleDataAdapter()
    '        _oda.InsertCommand = Ora_Trans_Command()
    '        _oda.UpdateCommand = Ora_Trans_Command()
    '        _oda.DeleteCommand = Ora_Trans_Command()
    '        Return _oda
    '    End Function
    '#End Region

#Region "eTracePublic"

    Public Function InvMigUserCheck(ByVal UserName As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim vCount As Integer
            Dim Sqlstr As String
            Dim ReturnValue As Boolean
            Try
                Sqlstr = String.Format(" select count(1) vCount from T_Config where ConfigID = 'INVMIGUSER' and Value like '%{0}%' ", UserName.ToUpper)
                vCount = Convert.ToDouble(da.ExecuteScalar(Sqlstr))
                If vCount = 0 Then
                    ReturnValue = False
                Else
                    ReturnValue = True
                End If

                Return ReturnValue
            Catch oe As Exception
                ErrorLogging("PublicFunction-InvMigUserCheck", "", oe.Message & oe.Source, "E")
                Return False
            End Try
        End Using
    End Function
    Public Function InvMigrationStatus(ByVal Status As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Dim Cvalue As Integer

                If Status = "Open" Then
                    Cvalue = 1
                ElseIf Status = "Close" Then
                    Cvalue = 0
                Else
                    Cvalue = -1
                    InvMigrationStatus = "The Status is blank"
                    Exit Function
                End If

                Sqlstr = String.Format("update T_Transaction set Active = {0} where TranID = 'QAL002' ", Cvalue)
                da.ExecuteScalar(Sqlstr)
                InvMigrationStatus = ""
            Catch ex As Exception
                ErrorLogging("PublicFunction-InvMigrationStatus", "", ex.Message & ex.Source, "E")
                InvMigrationStatus = ex.Message & ex.Source
            End Try
        End Using

    End Function
    Public Function InvMigCurrStatus() As Integer
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("select top 1 Active from T_Transaction where TranID = 'QAL002' ")
                InvMigCurrStatus = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
            Catch ex As Exception
                ErrorLogging("PublicFunction-InvMigCurrStatus", "", ex.Message & ex.Source, "E")
                InvMigCurrStatus = -1
            End Try
        End Using

    End Function
    Public Function GetMailList(ByVal eTraceModule As String) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetMailList  '{0}' ", eTraceModule)
                GetMailList = Convert.ToString(da.ExecuteScalar(Sqlstr))
            Catch ex As Exception
                ErrorLogging("PublicFunction-GetMailList", "", ex.Message & ex.Source, "E")
                GetMailList = ""
            End Try
        End Using
    End Function

    'Public Function ErrorLogging(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String) As Boolean
    '    Using da As DataAccess = GetDataAccess()
    '        Dim ErrFlag As Integer
    '        Try
    '            'Filter special character "'" from Error Message to avoid SP execution error
    '            ErrMsg = ErrMsg.Replace("'", "''")

    '            Dim Sqlstr As String
    '            Sqlstr = String.Format("exec sp_ErrorLogging '{0}', N'{1}', N'{2}'", ModuleName, User, ErrMsg)
    '            ErrFlag = da.ExecuteNonQuery(Sqlstr)
    '            If ErrFlag = -1 Then Return True
    '        Catch ex As Exception
    '            Return False
    '        End Try
    '    End Using

    'End Function
    Public Function ErrorLogging(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String, Optional ByVal Category As String = "") As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim ErrFlag As Integer
            Try
                'Filter special character "'" from Error Message to avoid SP execution error
                ErrMsg = ErrMsg.Replace("'", "''")

                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_ErrorLogging '{0}', N'{1}', N'{2}', N'{3}'", ModuleName, User, ErrMsg, Category)
                ErrFlag = da.ExecuteNonQuery(Sqlstr)

                If ErrFlag = -1 Then Return True
            Catch ex As Exception
                Return False
            End Try
        End Using

    End Function

    Public Function GetServerDate() As Date
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String

            Try
                strsql = String.Format("Select GetDate()")
                GetServerDate = da.ExecuteScalar(strsql)
            Catch ex As Exception
                ErrorLogging("PublicFunction-GetServerDate", "", ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function

    Public Function GetConfigValue(ByVal ConfigID As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetConfigValue  '{0}' ", ConfigID)
                GetConfigValue = Convert.ToString(da.ExecuteScalar(Sqlstr))

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetConfigValue", "", ex.Message & ex.Source, "E")
                GetConfigValue = ""
            End Try
        End Using

    End Function

    Public Function GetOrgServer(ByVal OrgCode As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select ServerName from T_Server with (nolock) where OrgCode = '{0}'", OrgCode)
                GetOrgServer = Convert.ToString(da.ExecuteScalar(Sqlstr))

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetOrgServer", "", "OrgCode: " & OrgCode & ex.Message & ex.Source, "E")
                GetOrgServer = ""
            End Try

        End Using
    End Function

    Public Function GetOrgID(ByVal OrgCode As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select OrgID from T_Server with (nolock) where OrgCode = '{0}'", OrgCode)
                GetOrgID = Convert.ToString(da.ExecuteScalar(Sqlstr))

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetOrgID", "", "OrgCode: " & OrgCode & ex.Message & ex.Source, "E")
                GetOrgID = ""
            End Try

        End Using
    End Function

    Public Function GetOrgLists(ByVal eTraceModule As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim OrgLists As DataSet = New DataSet

            Try
                Dim Sqlstr As String
                'Sqlstr = String.Format("Select OrgCode,OrgDescription from T_Server where ServerName = '{0}'", LoginServer)
                Sqlstr = String.Format("Select * from T_Server ")
                OrgLists = da.ExecuteDataSet(Sqlstr, "OrgCode")

                Dim ds As DataSet = New DataSet
                Dim ConfigID As String = "('ORGLIST', 'HHTIMEOUT', 'PCTIMEOUT', 'MSG001', 'HHF001', 'HHF002', 'HHF003', 'HHF004', 'CLID010', 'ACARD' )"
                Sqlstr = String.Format("Select ConfigID, Module, Value from T_Config with (nolock) where ConfigID in {0} ", ConfigID, eTraceModule)
                ds = da.ExecuteDataSet(Sqlstr, "Config")
                OrgLists.Merge(ds.Tables(0))

                Return OrgLists

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetOrgLists", "", "eTraceModule: " & eTraceModule & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using
    End Function

    Public Function GetLoginData(ByVal eTraceModule As String, ByVal TransactionID As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim LoginData As DataSet = New DataSet

            Try

                Dim Sqlstr As String
                'Sqlstr = String.Format("Select OrgCode,OrgDescription from T_Server where ServerName = '{0}'", LoginServer)
                Sqlstr = String.Format("Select * from T_Server ")
                LoginData = da.ExecuteDataSet(Sqlstr, "OrgCode")

                Dim ds As DataSet = New DataSet
                Dim ConfigID As String = "('ORGLIST', 'HHTIMEOUT', 'PCTIMEOUT', 'MSG001', 'HHF001', 'HHF002', 'HHF003', 'HHF004' )"
                If eTraceModule = "RECEIVING" Then
                    ConfigID = "('ORGLISTREC', 'HHTIMEOUT', 'PCTIMEOUT', 'MSG001', 'HHF001', 'HHF002', 'HHF003', 'HHF004' )"
                End If

                If eTraceModule = "CUSTOMREPORT" Then
                    ConfigID = "('ORGLISTCRPT', 'HHTIMEOUT', 'PCTIMEOUT', 'MSG001', 'HHF001', 'HHF002', 'HHF003', 'HHF004' )"
                End If

                'Sqlstr = String.Format("Select ConfigID, Module, Value from T_Config where ConfigID in {0} or (ConfigID like 'VER%' and Module = '{1}' )", ConfigID, eTraceModule)
                Sqlstr = String.Format("Select ConfigID, Module, Value from T_Config with (nolock) where ConfigID in {0} or (Module = '{1}' )", ConfigID, eTraceModule)
                ds = da.ExecuteDataSet(Sqlstr, "Config")
                If eTraceModule = "RECEIVING" Then
                    Dim DR() As DataRow
                    Dim i As Integer
                    DR = ds.Tables("Config").Select("ConfigID = 'ORGLISTREC'")
                    If DR.Length > 0 Then
                        For i = 0 To DR.Length - 1
                            If DR(i)("ConfigID") = "ORGLISTREC" Then
                                DR(i)("ConfigID") = "ORGLIST"
                            End If
                        Next
                        ds.Tables("Config").AcceptChanges()
                    End If
                End If

                If eTraceModule = "CUSTOMREPORT" Then
                    Dim DR() As DataRow
                    Dim i As Integer
                    DR = ds.Tables("Config").Select("ConfigID = 'ORGLISTCRPT'")
                    If DR.Length > 0 Then
                        For i = 0 To DR.Length - 1
                            If DR(i)("ConfigID") = "ORGLISTCRPT" Then
                                DR(i)("ConfigID") = "ORGLIST"
                            End If
                        Next
                        ds.Tables("Config").AcceptChanges()
                    End If
                End If

                LoginData.Merge(ds.Tables(0))

                If TransactionID <> "" Then
                    Dim PropVersion As String = ""
                    Sqlstr = String.Format("Select PropertyVersion from T_Transaction with (nolock) where TranID = '{0}' ", TransactionID)
                    PropVersion = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If PropVersion <> "" Then
                        Dim myDR As Data.DataRow
                        myDR = LoginData.Tables("Config").NewRow()
                        myDR("ConfigID") = "PROPVERSION"            ' "PropVersion"
                        myDR("Module") = TransactionID
                        myDR("Value") = PropVersion
                        LoginData.Tables("Config").Rows.Add(myDR)
                    End If
                End If

                Return LoginData

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetLoginData", "", "eTraceModule: " & eTraceModule & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using
    End Function

    Public Function GetHHVersion(ByVal ConfigID As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetConfigValue  '{0}' ", ConfigID)
                GetHHVersion = Convert.ToString(da.ExecuteScalar(Sqlstr))

            Catch ex As Exception
                ErrorLogging("PublicFunction-sp_GetHHVersion", "", "ConfigID: " & ConfigID & ex.Message & ex.Source, "E")
                GetHHVersion = "0"
            End Try

        End Using
    End Function

    Public Function GetGlobalRespID(ByVal OraInstance As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select GlobalRespID from T_GlobalRespID with (nolock) where OraInstance = '{0}'", OraInstance)
                GetGlobalRespID = Convert.ToString(da.ExecuteScalar(Sqlstr))

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetGlobalRespID", "", "OraInstance: " & OraInstance & ex.Message & ex.Source, "E")
                GetGlobalRespID = ""
            End Try

        End Using
    End Function

    Public Function GetScreenElements(ByVal GroupName As String, ByVal Lang As String, ByVal MessageClass As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            GetScreenElements = New DataSet
            Dim myConnect As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            Try

                Dim myDataTable As DataTable = New DataTable("ProcessHeader")
                Dim mySQLDA As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("SELECT a.ProcessID, convert(varchar(1),a.ProcessID) +  '. ' + b.Label as Name FROM T_ProcessHeader a, T_ProcessLang b where a.processid = b.processid and b.property = 'HEADER' and a.Groupname = @GroupName and b.Lang = @Lang ", myConnect)
                mySQLDA.SelectCommand.Parameters.AddWithValue("@GroupName", GroupName)
                mySQLDA.SelectCommand.Parameters.AddWithValue("@Lang", Lang)
                mySQLDA.Fill(myDataTable)
                GetScreenElements.Tables.Add(myDataTable)

                myDataTable = New DataTable("ProcessProperties")
                mySQLDA = New SqlClient.SqlDataAdapter("SELECT b.ProcessID, b.Property, c.Label, b.Required, b.DefaultValue,c.ValidationMessage, b.ValidationType, b.RegEx, b.DataType, b.FieldLength,b.NoOfDecimals,b.Visible,b.editable  FROM T_ProcessHeader a, T_ProcessProperties b, t_ProcessLang c where a.processid = b.processid and a.processid = c.processid and b.Property = c.Property and c.Lang = @Lang and a.Groupname = @GroupName ", myConnect)
                mySQLDA.SelectCommand.Parameters.AddWithValue("@GroupName", GroupName)
                mySQLDA.SelectCommand.Parameters.AddWithValue("@Lang", Lang)
                mySQLDA.Fill(myDataTable)
                GetScreenElements.Tables.Add(myDataTable)

                Dim Sqlstr As String
                If MessageClass = "REC" Then
                    Sqlstr = String.Format("Select MessageClass, MessageID, Message, Lang FROM T_Messages with (nolock) WHERE MessageClass ='REC' and Lang = @Lang ")
                Else
                    Sqlstr = String.Format("Select MessageClass, MessageID, Message, Lang FROM T_Messages with (nolock) WHERE Lang = @Lang ")
                End If

                myDataTable = New DataTable("Messages")
                mySQLDA = New SqlClient.SqlDataAdapter(Sqlstr, myConnect)
                mySQLDA.SelectCommand.Parameters.AddWithValue("@Lang", Lang)
                mySQLDA.Fill(myDataTable)
                GetScreenElements.Tables.Add(myDataTable)

                Dim ds As DataSet = New DataSet
                Sqlstr = String.Format("Select ConfigID, Module, Value from T_Config with (nolock) where (ConfigID like 'REC%' or ConfigID like 'VER%') and Module = '{0}' ", GroupName)
                ds = da.ExecuteDataSet(Sqlstr, "PropVersion")
                GetScreenElements.Merge(ds.Tables(0))
            Catch ex As Exception
            End Try

        End Using

    End Function

    Public Function WritePrintLabel(ByVal LabelSeqNo As String, ByVal LabelFile As String, ByVal LabelPrinter As String, ByVal Content As String)
        Dim myCMD As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strCMD As String
        strCMD = String.Format("insert into T_LabelPrint values ('{0}','{1}','{2}','{3}',getdate(),'','1')", LabelSeqNo, LabelFile, LabelPrinter, Content)
        Try
            myConn.Open()
            myCMD = New SqlClient.SqlCommand(strCMD, myConn)
            myCMD.CommandTimeout = TimeOut_M5
            myCMD.ExecuteNonQuery()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-WritePrintLabel", "", "LabelSeqNo: " & LabelSeqNo & ", " & ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return ""

    End Function

    Public Function OpenPrintLabel() As DataSet
        OpenPrintLabel = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabel.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabel.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
                    & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
                    & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1' and LabelPrinter not like 'TEMP%'  " _
                    & " and LabelFile in ('CLMatLabel.lab','POLabel.lab','RTLabel.lab','B1B_Carton.lab', 'CartonLabel.lab', 'PickingLabel.lab', 'SPGLabel.lab', 'ECSMTPickLabel.lab', 'LDPickingLabel.lab', 'LDSpareLabel.lab', 'ProcessLabel.lab', 'CH09Label.lab') " _
                    & " GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
                    & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 20 ORDER BY CreatedOn"

        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 order by Content", myConn)

        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function OpenPrintLabel_Production() As DataSet
        OpenPrintLabel_Production = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabel_Production.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabel_Production.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
                    & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
                    & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM T_LabelPrint with (nolock) WHERE StatusCode = '1' and LabelPrinter not like 'TEMP%'  " _
                    & " and LabelFile not in ('CLMatLabel.lab','POLabel.lab','RTLabel.lab','B1B_Carton.lab', 'CartonLabel.lab', 'PickingLabel.lab', 'SPGLabel.lab', 'ECSMTPickLabel.lab', 'LDPickingLabel.lab', 'LDSpareLabel.lab', 'ProcessLabel.lab', 'CH09Label.lab') " _
                    & " GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
                    & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 20 ORDER BY CreatedOn"

        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 order by Content", myConn)

        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel_Production", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function OpenPrintLabel_LAG() As DataSet
        OpenPrintLabel_LAG = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabel_LAG.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabel_LAG.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
                    & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
                    & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM T_LabelPrint with (nolock) WHERE StatusCode = '1' GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
                    & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 10 ORDER BY CreatedOn"

        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 order by Content", myConn)

        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                If Not drHeader.GetValue(2) Is DBNull.Value Then
                    If Not Left(drHeader.GetValue(2).ToString.ToUpper, 5) = "PHCVT" Then
                        mydatarow = Labels.NewRow()
                        If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                        If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                        If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                        If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                        If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                        If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                        Labels.Rows.Add(mydatarow)
                        If Not drHeader.GetValue(3) Is DBNull.Value Then
                            arryDetail = Split(mydatarow("Content").ToString, "^")
                            For i = 0 To UBound(arryDetail) Step 2
                                mydatarow = Contents.NewRow()
                                mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                                mydatarow("title") = arryDetail(i)
                                mydatarow("detail") = arryDetail(i + 1)
                                Contents.Rows.Add(mydatarow)
                            Next
                        End If
                    End If
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel_LAG", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function OpenPrintLabel_CVT() As DataSet
        OpenPrintLabel_CVT = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabel_CVT.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabel_CVT.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
                    & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
                    & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM T_LabelPrint with (nolock) WHERE StatusCode = '1' GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
                    & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 10 ORDER BY CreatedOn"

        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 order by Content", myConn)

        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                If Not drHeader.GetValue(2) Is DBNull.Value Then
                    If Left(drHeader.GetValue(2).ToString.ToUpper, 5) = "PHCVT" Then
                        mydatarow = Labels.NewRow()
                        If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                        If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                        If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                        If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                        If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                        If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                        Labels.Rows.Add(mydatarow)
                        If Not drHeader.GetValue(3) Is DBNull.Value Then
                            arryDetail = Split(mydatarow("Content").ToString, "^")
                            For i = 0 To UBound(arryDetail) Step 2
                                mydatarow = Contents.NewRow()
                                mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                                mydatarow("title") = arryDetail(i)
                                mydatarow("detail") = arryDetail(i + 1)
                                Contents.Rows.Add(mydatarow)
                            Next
                        End If
                    End If
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel_CVT", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function OpenPrintLabelForApps() As DataSet
        OpenPrintLabelForApps = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForApps.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForApps.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode from T_LabelPrint with (nolock) where statusCode = 1 and (LabelPrinter = 'PHLAGINT01' or LabelPrinter = 'PHLAGINT02' or LabelPrinter = 'PHLAGINT03' or LabelPrinter = 'PHLAGINT04' or LabelPrinter = 'PHLAGINT05' or LabelPrinter = 'PHLAGINT06' or LabelPrinter = 'PHLAGINT07' or " _
                & "LabelPrinter = 'PHLAGWRL03' or LabelPrinter = 'PHLAGWRL04' or LabelPrinter = 'PHLAGWRL05' or LabelPrinter = 'PHLAGWRL06' or LabelPrinter = 'PHLAGWRL07' or " _
                & "LabelPrinter = 'PHLAGWRL08' or LabelPrinter = 'PHLAGWRL09' or LabelPrinter = 'PHLAGWRL10' or LabelPrinter = 'PHLAGWRL11' or LabelPrinter = 'PHLAGWRL12' or LabelPrinter = 'PHLAGWRL13' or LabelPrinter = 'PHLAGWRL14' or " _
                & "LabelPrinter = 'PHLAGWRL15' or LabelPrinter = 'PHLAGWRL16' or LabelPrinter = 'PHLAGWRL17' or LabelPrinter = 'PHLAGWRL18' or LabelPrinter = 'PHLAGWRL19' or LabelPrinter = 'PHLAGWRL20' or LabelPrinter = 'PHLAGWRL21' or LabelPrinter = 'PHLAGWRL22' or LabelPrinter = 'PHLAGWRL23') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function OpenPrintLabelForEtrace2() As DataSet
        OpenPrintLabelForEtrace2 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForEtrace2.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForEtrace2.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode from T_LabelPrint with (nolock) where statusCode = 1 and (LabelPrinter = 'PHLAGINT08' or LabelPrinter = 'PHLAGINT09' or LabelPrinter = 'PHLAGWRL01' or LabelPrinter = 'PHLAGWRL02' or LabelPrinter = 'PHLAGZEB01' or LabelPrinter = 'PHLAGPRT01' or LabelPrinter = 'PHLAGPRT02' or " _
                & "LabelPrinter = 'PHLAGINT10' or LabelPrinter = 'PHLAGINT11' or LabelPrinter = 'PHLAGWRL21' or LabelPrinter = 'PHLAGWRL22' or LabelPrinter = 'PHLAGWRL23' or LabelPrinter = 'PHLAGWRL24' or LabelPrinter = 'PHLAGWRL25' or " _
                & "LabelPrinter = 'PHLAGWRL26' or LabelPrinter = 'PHLAGWRL27' or LabelPrinter = 'PHLAGWRL28' or LabelPrinter = 'PHLAGWRL29' or LabelPrinter = 'PHLAGWRL30' or LabelPrinter = 'PHLAGWRL31' or LabelPrinter = 'PHLAGWRL32' or " _
                & "LabelPrinter = 'PHLAGWRL33' or LabelPrinter = 'PHLAGWRL34' or LabelPrinter = 'PHLAGWRL35' or LabelPrinter = 'PHLAGWRL36' or LabelPrinter = 'PHLAGWRL37' or LabelPrinter = 'PHLAGWRL38' or LabelPrinter = 'PHLAGWRL39' or " _
                & "LabelPrinter = 'PHLAGWRL40' or LabelPrinter = 'PHLAGWRL41') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function OpenPrintLabelForPrtSvr() As DataSet
        OpenPrintLabelForPrtSvr = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForPrtSvr.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForPrtSvr.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode from T_LabelPrint with (nolock) where statusCode = 1 and (LabelPrinter = 'PHCVTINT01' or LabelPrinter = 'PHCVTINT02' or LabelPrinter = 'PHCVTINT03' or LabelPrinter = 'PHCVTPRT01' OR " _
                & "LabelPrinter = 'PHCVTINT04' or LabelPrinter = 'PHCVTINT05' or LabelPrinter = 'PHCVTINT06' or LabelPrinter = 'PHCVTINT07' or LabelPrinter = 'PHCVTWRL01' or LabelPrinter = 'PHCVTWRL02' or LabelPrinter = 'PHCVTWRL03' or LabelPrinter = 'PHCVTWRL04' or " _
                & "LabelPrinter = 'PHCVTWRL05' or LabelPrinter = 'PHCVTWRL06' or LabelPrinter = 'PHCVTWRL07' or LabelPrinter = 'PHCVTWRL08' or LabelPrinter = 'PHCVTWRL09' or LabelPrinter = 'PHCVTWRL10' or LabelPrinter = 'PHCVTWRL11' or " _
                & "LabelPrinter = 'PHCVTWRL12' or LabelPrinter = 'PHCVTWRL13' or LabelPrinter = 'PHCVTWRL14' or LabelPrinter = 'PHCVTWRL15' or LabelPrinter = 'PHCVTWRL16' or LabelPrinter = 'PHCVTWRL17' or LabelPrinter = 'PHCVTWRL18' or " _
                & "LabelPrinter = 'PHCVTWRL19' or LabelPrinter = 'PHCVTWRL20' or LabelPrinter = 'PHCVTWRL21' or LabelPrinter = 'PHCVTWRL22' or LabelPrinter = 'PHCVTWRL23' or LabelPrinter = 'PHCVTWRL24') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function


    Public Function OpenPrintLabelForTemp01() As DataSet
        OpenPrintLabelForTemp01 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp01.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp01.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 and (LabelPrinter = 'TEMP01001' or LabelPrinter = 'TEMP01002') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function OpenPrintLabelForTemp02() As DataSet
        OpenPrintLabelForTemp02 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp02.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp02.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 and (LabelPrinter = 'TEMP02001' or LabelPrinter = 'TEMP02002') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function OpenPrintLabelForTemp03() As DataSet
        OpenPrintLabelForTemp03 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp03.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp03.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 and (LabelPrinter = 'TEMP03001' or LabelPrinter = 'TEMP03002') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function OpenPrintLabelForTemp04() As DataSet
        OpenPrintLabelForTemp04 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp04.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp04.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 and (LabelPrinter = 'TEMP04001' or LabelPrinter = 'TEMP04002') order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function


    Public Function OpenPrintLabelForTemp05() As DataSet
        OpenPrintLabelForTemp05 = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp05.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintLabelForTemp05.Tables.Add(Contents)

        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header

        'Updated on Oct 14, 2010
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE statusCode = 1 and (LabelPrinter = 'TEMP05001' or LabelPrinter = 'TEMP05002')  order by Content"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        Try
            myConn.Open()
            cmdReadHeader.CommandTimeout = TimeOut_M5
            drHeader = cmdReadHeader.ExecuteReader()

            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function updatePrintLabel(ByVal SeqNo As String, ByVal ErrMsg As String)
        Dim myCMD As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strCMD As String

        Try
            If ErrMsg = "" Then
                strCMD = String.Format("UPDATE T_LabelPrint set PrintedOn = getdate(), statusCode='0' where LabelSeqNo='{0}'", SeqNo)
            Else
                strCMD = String.Format("UPDATE T_LabelPrint set statusCode='0',ErrMsg='{0}' where LabelSeqNo='{1}'", ErrMsg, SeqNo)
            End If

            myConn.Open()
            myCMD = New SqlClient.SqlCommand(strCMD, myConn)
            myCMD.CommandTimeout = TimeOut_M5
            myCMD.ExecuteNonQuery()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-updatePrintLabel", "", "LabelSeqNo: " & SeqNo & ", " & ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return ""

    End Function


    Public Function PrintCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean

        PrintCLIDs = True
        Try
            If CLIDs Is Nothing OrElse CLIDs.Tables.Count = 0 OrElse CLIDs.Tables(0).Rows.Count = 0 Then
                PrintCLIDs = False
                Exit Function
            End If

            'Sort CLID by Ascending before written to eTrace table T_LabelPrint, this is only for LD / PH site           -- 2/13/2016
            'Please note ZS will NOT sort CLID here, so this already comment out in ZS Web Service                            -- 2/13/2016
            'Dim dtPrint As DataTable = New DataTable
            'Dim SortColName As String = CLIDs.Tables(0).Columns(0).ColumnName
            'SortColName = SortColName & " ASC"

            'CLIDs.Tables(0).DefaultView.Sort = SortColName
            'dtPrint = CLIDs.Tables(0).DefaultView.ToTable()

            'CLIDs = New DataSet
            'CLIDs.Tables.Add(dtPrint)	    

            Dim i As Integer
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                If CLIDs.Tables(0).Rows(i)(0).ToString <> "" Then
                    If PrintCLID(CLIDs.Tables(0).Rows(i)(0), Printer) = False Then
                        PrintCLIDs = False
                    End If
                Else
                    ErrorLogging("PublicFunction-PrintCLIDs", "", "CLID not found ", "I")
                    PrintCLIDs = False
                End If
                Sleep(5)
            Next
        Catch ex As Exception
            ErrorLogging("PublicFunction-PrintCLIDs", "", ex.Message & ex.Source, "E")
            PrintCLIDs = False
        End Try

    End Function

    ''' <summary>
    ''' 
    ''' Print CLID function.
    ''' 
    ''' </summary>
    ''' <param name="CLID"></param>
    ''' <param name="Printer"></param>
    ''' <returns></returns>
    ''' <remarks> 
    ''' Add Lot No field in PCB Lable 06-29-2011.
    ''' Add Date Code in PCB Lable
    ''' </remarks>
    ''' <modify> By Jackson Huang</modify>
    ''' <Date>12-29-2011 </Date>
    Public Function PrintCLID(ByVal CLID As String, ByVal Printer As String) As Boolean
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim lblData As MatLabel = New MatLabel
        Dim LabelPrinter, lblPrint, LabelType, StartStr, MidStr As String
        Dim TmpDate As Date
        Dim PCBLabel As POLabel = New POLabel

        LabelPrinter = Printer
        PrintCLID = True

        '0  CLID, OrgCode, MaterialNo, MaterialRevision, QtyBaseUOM, BaseUOM, LotNo, DateCode, ExpDate, RecDate, RecDocNo, RoHS, StockType, '12
        '13 MatSuffix1, MatSuffix2, MatSuffix3, StatusCode, AddlData, Manufacturer, ManufacturerPN, Stemp, MSL '21

        Dim LblStart As String = Microsoft.VisualBasic.Left(CLID, 2)                     'LD Huawei requirement   -- 11/16/2016

        StartStr = Microsoft.VisualBasic.Left(CLID, 1)
        MidStr = Microsoft.VisualBasic.Mid(CLID, 3, 1)

        If StartStr = "B" OrElse MidStr = "P" OrElse LblStart = "LE" Then
            LabelType = "PCBID"
            CLMasterSQLCommand = New SqlClient.SqlCommand("Select CLID,OrgCode,MaterialNo,MaterialRevision,PurOrdNo,QtyBaseUOM,BaseUOM,RecDate,StatusCode,SONo,SOLine, LotNo, DateCode, ManufacturerPN, RoHS from T_CLMaster with (nolock) where CLID = @CLID ", myConn)
        ElseIf MidStr = "B" Then
            CLMasterSQLCommand = New SqlClient.SqlCommand("Select BoxID,OrgCode,MaterialNo,MaterialRevision,Sum(QtyBaseUOM) As BoxQty,BaseUOM,'' as LotNo,'' as DateCode,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot,'' as MCPosition, '' as PredefinedSubInv, '' as ItemText, '' as COO from T_CLMaster with (nolock) group by BoxID,OrgCode,MaterialNo,MaterialRevision,BaseUOM,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot having StatusCode > 0 AND BoxID = @CLID ", myConn)
            lblData.LabelType = "Box Label"
            LabelType = "BoxID"
        ElseIf IsNumeric(MidStr) = True OrElse MidStr = "V" Then
            CLMasterSQLCommand = New SqlClient.SqlCommand("Select CLID,OrgCode,MaterialNo,MaterialRevision,QtyBaseUOM,BaseUOM,LotNo,DateCode,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot,MCPosition, PredefinedSubInv, ItemText, CountryOfOrigin as COO from T_CLMaster with (nolock) where CLID = @CLID ", myConn)
            lblData.LabelType = "Material Label"
            LabelType = "CLID"
        Else
            ErrorLogging("PublicFunction-PrintCLID", "", "CLID: " & CLID & " not in the label type list ", "I")
            PrintCLID = False
            Exit Function
        End If

        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters("@CLID").Value = CLID

        Try
            myConn.Open()
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                If LabelType = "PCBID" Then
                    'If objReader.GetValue(8) = "0" Then       ' Will not print any scrapped PCB
                    '    PrintCLID = False
                    '    Exit Function
                    'End If

                    ' Record PCB Label Info
                    If Not objReader.GetValue(0) Is DBNull.Value Then PCBLabel.CLID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then PCBLabel.OrgCode = objReader.GetValue(1) 'OrgCode
                    If Not objReader.GetValue(2) Is DBNull.Value Then PCBLabel.PCBA = objReader.GetValue(2)
                    If Not objReader.GetValue(3) Is DBNull.Value Then PCBLabel.Rev = objReader.GetValue(3)
                    If Not objReader.GetValue(4) Is DBNull.Value Then PCBLabel.POID = objReader.GetValue(4)

                    'Dim POQty As Integer = 0
                    'If Not objReader.GetValue(5) Is DBNull.Value Then POQty = CInt(objReader.GetValue(5))
                    Dim POQty As Decimal = 0
                    If Not objReader.GetValue(5) Is DBNull.Value Then POQty = objReader.GetValue(5)

                    PCBLabel.Qty = Format(POQty, "#0.#####")

                    If Not objReader.GetValue(6) Is DBNull.Value Then PCBLabel.UOM = objReader.GetValue(6)
                    If Not objReader.GetValue(7) Is DBNull.Value Then
                        TmpDate = objReader.GetValue(7)
                        PCBLabel.RecDate = TmpDate.ToString("MM/dd/yyyy")
                    End If

                    If Not objReader.GetValue(9) Is DBNull.Value Then PCBLabel.SONo = objReader.GetValue(9)
                    If Not objReader.GetValue(10) Is DBNull.Value Then PCBLabel.SOLine = objReader.GetValue(10)
                    If Not objReader.GetValue(11) Is DBNull.Value Then PCBLabel.LotNo = objReader.GetValue(11) ''Add Lot No to PCB Lable.

                    '' Add DateCode to PCBA Lable.
                    If Not objReader.GetValue(12) Is DBNull.Value Then PCBLabel.DateCode = objReader.GetValue(12)

                    'Add MPN to PCBA Label for Mix Source Build Project   09/09/2013            
                    If Not objReader.GetValue(13) Is DBNull.Value Then PCBLabel.MPN = objReader.GetValue(13)

                    'Add RoHS to PCBA Label    05/19/2015            
                    If Not objReader.GetValue(14) Is DBNull.Value Then PCBLabel.RoHS = objReader.GetValue(14)

                    lblPrint = PrintPOLabel(LabelPrinter, PCBLabel)
                Else

                    'If objReader.GetValue(16) = "0" Then      '   Will not print any label for cancelled GR
                    '    PrintCLID = False
                    '    Exit Function
                    'End If
                    If Not objReader.GetValue(0) Is DBNull.Value Then lblData.LabelID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then lblData.OrgCode = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then lblData.Material = objReader.GetValue(2)
                    If Not objReader.GetValue(3) Is DBNull.Value Then lblData.Rev = objReader.GetValue(3)
                    If Not objReader.GetValue(4) Is DBNull.Value Then lblData.Qty = objReader.GetValue(4)
                    If Not objReader.GetValue(5) Is DBNull.Value Then lblData.UOM = objReader.GetValue(5)
                    If Not objReader.GetValue(6) Is DBNull.Value Then lblData.LotNo = objReader.GetValue(6)
                    If Not objReader.GetValue(7) Is DBNull.Value Then lblData.DCode = objReader.GetValue(7)

                    lblData.Qty = Format(objReader.GetValue(4), "#0.#####")
                    'Dim LabelQty As Double = Val(objReader.GetValue(4))
                    'lblData.Qty = Math.Round(LabelQty, 5).ToString

                    If Not objReader.GetValue(8) Is DBNull.Value Then
                        TmpDate = objReader.GetValue(8)
                        lblData.ExpDate = TmpDate.ToString("MM/dd/yyyy")
                    End If
                    If Not objReader.GetValue(9) Is DBNull.Value Then
                        TmpDate = objReader.GetValue(9)
                        lblData.RecDate = TmpDate.ToString("MM/dd/yyyy")
                    End If

                    'Default set to print out Barcode for Material
                    lblData.IPN = lblData.Material

                    If Not objReader.GetValue(10) Is DBNull.Value Then lblData.RT = objReader.GetValue(10)
                    If Not objReader.GetValue(11) Is DBNull.Value Then lblData.RoHS = objReader.GetValue(11)
                    If Not objReader.GetValue(12) Is DBNull.Value Then lblData.InspFlag = objReader.GetValue(12)
                    'If Not objReader.GetValue(17) Is DBNull.Value Then lblData.Safety = objReader.GetValue(17)
                    If Not objReader.GetValue(18) Is DBNull.Value Then lblData.Mfr = objReader.GetValue(18)
                    If Not objReader.GetValue(19) Is DBNull.Value Then lblData.MfrPN = objReader.GetValue(19)
                    If Not objReader.GetValue(20) Is DBNull.Value Then lblData.Stemp = objReader.GetValue(20)
                    If Not objReader.GetValue(21) Is DBNull.Value Then lblData.MSL = objReader.GetValue(21)

                    'Check if RTLot is different from RT, if yes, then Combine RT / RTLot into field lblData.RT
                    If objReader.GetValue(22) Is DBNull.Value OrElse Trim(objReader.GetValue(22)) = "" Then
                    Else
                        Dim RT As String = lblData.RT
                        Dim RTLot As String = Trim(objReader.GetValue(22))
                        'If RTLot <> RT Then lblData.RT = lblData.RT & " / " & RTLot
                        If RTLot <> RT Then                                                 'RTLot generated by RT-ExpDate (yyMMdd)
                            If RTLot.Contains("-") Then
                                lblData.RT = RTLot
                            Else
                                lblData.RT = lblData.RT & " / " & RTLot
                            End If
                        End If
                    End If

                    'Print MCPosition on MaterialLabel requested by ZS only  --- 06/24/2014
                    If Not objReader.GetValue(23) Is DBNull.Value Then lblData.MCPosition = objReader.GetValue(23)

                    'Print PredefinedSubInv on MaterialLabel requested by ZS only  --- 03/27/2015
                    If Not objReader.GetValue(24) Is DBNull.Value Then lblData.PreSubInv = objReader.GetValue(24)

                    'Print ItemText on MaterialLabel requested by LD Huawei project  --- 12/27/2016
                    If Not objReader.GetValue(25) Is DBNull.Value Then lblData.ItemText = objReader.GetValue(25)

                    'Print COO on MaterialLabel requested by ZS only  --- 12/27/2016
                    If Not objReader.GetValue(26) Is DBNull.Value Then lblData.COO = objReader.GetValue(26)

                    lblData.Safety = ""
                    lblData.ESD = BlankESD
                    If Not objReader.GetValue(17) Is DBNull.Value Then
                        Dim AddlData As String = objReader.GetValue(17).ToString.ToUpper
                        If InStr(AddlData, "S") > 0 Then lblData.Safety = "Safety"
                        If InStr(AddlData, "E") > 0 Then lblData.ESD = ESDImage
                    End If


                    Dim flag As Integer
                    flag = 0

                    If objReader.GetValue(13) Is DBNull.Value OrElse Trim(objReader.GetValue(13)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & objReader.GetValue(13) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If
                    If objReader.GetValue(14) Is DBNull.Value OrElse Trim(objReader.GetValue(14)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & objReader.GetValue(14) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If
                    If objReader.GetValue(15) Is DBNull.Value OrElse Trim(objReader.GetValue(15)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & objReader.GetValue(15) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If
                    If flag = 0 Then
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                    End If

                    If lblPrint <> "True" Then
                        PrintCLID = False
                    End If
                End If
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-PrintCLID", "", "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
            PrintCLID = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function PrintLabel(ByVal Printer As String, ByVal LabelData As MatLabel, Optional ByVal NoOfLabels As Integer = 1) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim strContent As String
            Dim arryFile() As String
            Try
                arryFile = Split(CLIDLabelFile, "\")
                LabelData.MfrPN = Replace(LabelData.MfrPN, "^", "~")
                strContent = "LabelID^" & LabelData.LabelID & "^OrgCode^" & LabelData.OrgCode & "^RT^" & LabelData.RT & "^Material^" & LabelData.Material _
                                        & "^Rev^" & LabelData.Rev & "^Qty^" & LabelData.Qty & "^UOM^" & LabelData.UOM & "^ExpDate^" & LabelData.ExpDate & "^RoHS^" & LabelData.RoHS _
                                        & "^DCode^" & LabelData.DCode & "^LotNo^" & LabelData.LotNo & "^InspFlag^" & LabelData.InspFlag & "^RecDate^" & LabelData.RecDate _
                                        & "^MatSuffix^" & LabelData.MatSuffix & "^LabelType^" & LabelData.LabelType & "^Safety^" & LabelData.Safety & "^ESD^" & LabelData.ESD _
                                        & "^Mfr^" & LabelData.Mfr & "^MfrPN^" & LabelData.MfrPN & "^Stemp^" & LabelData.Stemp & "^MSL^" & LabelData.MSL & "^MCPosition^" & LabelData.MCPosition _
                                        & "^PreSubInv^" & LabelData.PreSubInv & "^ItemText^" & LabelData.ItemText & "^COO^" & LabelData.COO & "^PNSeq^" & LabelData.PNSeq & "^IPN^" & LabelData.IPN

                sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                da.ExecuteScalar(sql)
                PrintLabel = True
            Catch ex As Exception
                ErrorLogging("PublicFunction-PrintLabel", "", "LabelID: " & LabelData.LabelID & ", " & ex.Message & ex.Source, "E")
                PrintLabel = ex.Message & ex.Source
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Print Label Function.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks> 
    ''' Add LotNo in PCBA label 06-29-2011.
    ''' Add Date Code in PCBA label.
    ''' </remarks>
    ''' <modify> By Jackson Huang </modify>
    ''' <Date>12-28-2011</Date>
    ''' 

    Public Function PrintPOLabel(ByVal Printer As String, ByVal LabelData As POLabel) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim strContent As String
            Dim arryFile() As String
            Try
                arryFile = Split(POLabelFile, "\")
                strContent = "CLID^" & LabelData.CLID & "^OrgCode^" & LabelData.OrgCode & "^PCBA^" & LabelData.PCBA & "^POID^" & LabelData.POID & "^Qty^" & LabelData.Qty & "^UOM^" & LabelData.UOM & "^Rev^" & LabelData.Rev _
                           & "^RecDate^" & LabelData.RecDate & "^SONo^" & LabelData.SONo & "^SOLine^" & LabelData.SOLine & "^LotNo^" & LabelData.LotNo & "^DateCode^" & LabelData.DateCode & "^MPN^" & LabelData.MPN & "^RoHS^" & LabelData.RoHS

                sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                da.ExecuteScalar(sql)
                PrintPOLabel = True
            Catch ex As Exception
                ErrorLogging("PublicFunction-PrintPOLabel", "", "POLabelID: " & LabelData.CLID & ", " & ex.Message & ex.Source, "E")
                PrintPOLabel = ex.Message & ex.Source
            End Try

        End Using
    End Function

    'Public Function GetProcessProperties(ByVal GroupName As String, ByVal Lang As String, ByVal ClientVersion As String) As DataSet

    '    GetProcessProperties = New DataSet

    '    If Not ClientVersion = "" Then
    '        Dim mySQLCommand As SqlClient.SqlCommand
    '        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
    '        Dim objReader As SqlClient.SqlDataReader
    '        Dim ServerVersion As String = ""

    '        'Read property version
    '        mySQLCommand = New SqlClient.SqlCommand("Select Message from T_Messages where MessageClass = @MessageClass and MessageID = @MessageID and Lang = @Lang", myConn)
    '        mySQLCommand.Parameters.Add("@MessageClass", SqlDbType.VarChar, 3, "MessageClass")
    '        mySQLCommand.Parameters.Add("@MessageID", SqlDbType.Int, 3, "MessageID")
    '        mySQLCommand.Parameters.Add("@Lang", SqlDbType.VarChar, 50, "Lang")
    '        mySQLCommand.Parameters("@MessageClass").Value = "APP"
    '        mySQLCommand.Parameters("@MessageID").Value = 999
    '        mySQLCommand.Parameters("@Lang").Value = "ENG"

    '        Try
    '            myConn.Open()
    '            mySQLCommand.CommandTimeout = TimeOut_M5
    '            objReader = mySQLCommand.ExecuteReader()
    '            While objReader.Read()
    '                If Not objReader.GetValue(0) Is DBNull.Value Then ServerVersion = objReader.GetValue(0)
    '            End While
    '            myConn.Close()
    '        Catch ex As Exception
    '        Finally
    '            If myConn.State <> ConnectionState.Closed Then myConn.Close()
    '        End Try

    '        'If the version is the same, then no update is required.
    '        If ServerVersion = ClientVersion Then
    '            Exit Function
    '        End If
    '    End If

    '    Dim MessageClass As String = "APP"
    '    GetProcessProperties = GetScreenElements(GroupName, Lang, MessageClass)
    'End Function
    Public Function GetProcessProperties(ByVal GroupName As String, ByVal Lang As String, ByVal TransactionID As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim Msg As String = "GroupName/TransactionID: " & GroupName & "/ " & TransactionID & "; "

            Dim Sqlstr As String
            GetProcessProperties = New DataSet

            Try
                Sqlstr = String.Format("exec sp_GetProcessProperties  '{0}','{1}','{2}' ", GroupName, Lang, TransactionID)

                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"ProcessHeader", "ProcessProperties", "Messages", "PropVersion"}
                GetProcessProperties = da.ExecuteDataSet(sql, tables)
            Catch ex As Exception
                ErrorLogging("PublicFunction-GetProcessProperties", "", Msg & ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function

    Public Function GetAccessCardUserInfo(ByVal AccessCardID As String) As AccessCard
        GetAccessCardUserInfo = New AccessCard

        Dim ACard As String = ""

        Try
            Dim Sqlstr As String
            Dim da As DataAccess = GetDataAccess()
            Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID ='ACARD' ")
            ACard = Convert.ToString(da.ExecuteScalar(Sqlstr))
        Catch ex As Exception
            ErrorLogging("PublicFunction-GetAccessCardUserInfo1", "", AccessCardID & ", " & ex.Message & ex.Source, "E")
        End Try

        If ACard = "NO" Then
            Try
                Dim CourseCode As String = "3CP44"              'For FY / LD eTrace Access Card Login
                Return GetAccessCardUserInfoCourse(AccessCardID, CourseCode)
            Catch ex As Exception
                ErrorLogging("PublicFunction-GetAccessCardUserInfo2", "", AccessCardID & ", " & ex.Message & ex.Source, "E")
            End Try
            Exit Function
        End If


        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("HRISDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        Dim strHex, strHex1, strHex2 As String
        Dim IntDec1, IntDec2, EmpCardID As Integer

        Try

            strHex = Hex(AccessCardID).PadLeft(7, "0")
            strHex1 = Left(Right(strHex, 7), 3)
            strHex2 = Right(strHex, 4)

            IntDec1 = Convert.ToInt32(strHex1, 16)
            IntDec2 = Convert.ToInt32(strHex2, 16)

            EmpCardID = Val(Str(IntDec1) & IntDec2.ToString.PadLeft(5, "0"))

            AccessCardID = EmpCardID
            CLMasterSQLCommand = New SqlClient.SqlCommand("Select AccessCardID, EmployeeID,Name,Dept from T_AccessCard where AccessCardID = @EmpCardID", myConn)
            CLMasterSQLCommand.Parameters.Add("@EmpCardID", SqlDbType.VarChar, 30, "EmpCardID")
            CLMasterSQLCommand.Parameters("@EmpCardID").Value = AccessCardID

            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            If Not objReader.HasRows Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("Select AccessCardID, EmployeeID,Name,Dept from T_AccessCardBackup where AccessCardID = @EmpCardID", myConn)
                CLMasterSQLCommand.Parameters.Add("@EmpCardID", SqlDbType.VarChar, 30, "EmpCardID")
                CLMasterSQLCommand.Parameters("@EmpCardID").Value = AccessCardID
                objReader.Close()
                objReader = CLMasterSQLCommand.ExecuteReader()
                While objReader.Read()
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetAccessCardUserInfo.AccessCardID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetAccessCardUserInfo.EmplogeeID = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfo.CHNName = objReader.GetValue(2)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfo.Name = objReader.GetValue(1) 'Use EmplogeeID insteads of UserName
                    If Not objReader.GetValue(3) Is DBNull.Value Then GetAccessCardUserInfo.Dept = objReader.GetValue(3)
                End While
            Else
                While objReader.Read()
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetAccessCardUserInfo.AccessCardID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetAccessCardUserInfo.EmplogeeID = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfo.CHNName = objReader.GetValue(2)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfo.Name = objReader.GetValue(1) 'Use EmplogeeID insteads of UserName
                    If Not objReader.GetValue(3) Is DBNull.Value Then GetAccessCardUserInfo.Dept = objReader.GetValue(3)
                End While
            End If
        Catch ex As Exception
            ErrorLogging("PublicFunction-GetAccessCardUserInfo", "", AccessCardID & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function GetAccessCardUserInfoCourse(ByVal AccessCardID As String, ByVal CourseCode As String) As AccessCard
        GetAccessCardUserInfoCourse = New AccessCard

        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("HRISDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        'Dim strHex, strHex1, strHex2 As String
        'Dim IntDec1, IntDec2, EmpCardID As Integer

        Try

            'strHex = Hex(AccessCardID).PadLeft(7, "0")
            'strHex1 = Left(Right(strHex, 7), 3)
            'strHex2 = Right(strHex, 4)

            'IntDec1 = Convert.ToInt32(strHex1, 16)
            'IntDec2 = Convert.ToInt32(strHex2, 16)

            'EmpCardID = Val(Str(IntDec1) & IntDec2.ToString.PadLeft(5, "0"))

            CLMasterSQLCommand = New SqlClient.SqlCommand("Select AccessCardID, EmployeeID,Name from T_AccessCard where AccessCardID = '" + AccessCardID + "' and CertifiedCourse like '%" + CourseCode + "%'", myConn)
            'CLMasterSQLCommand.Parameters.Add("@EmpCardID", SqlDbType.VarChar, 30, "EmpCardID")
            'CLMasterSQLCommand.Parameters("@EmpCardID").Value = AccessCardID

            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            If Not objReader.HasRows Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("Select AccessCardID, EmployeeID,Name from T_AccessCardBackup where AccessCardID = '" + AccessCardID + "' and CertifiedCourse like '%" + CourseCode + "%'", myConn)
                'CLMasterSQLCommand.Parameters.Add("@EmpCardID", SqlDbType.VarChar, 30, "EmpCardID")
                'CLMasterSQLCommand.Parameters("@EmpCardID").Value = AccessCardID
                objReader.Close()
                objReader = CLMasterSQLCommand.ExecuteReader()
                While objReader.Read()
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetAccessCardUserInfoCourse.AccessCardID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetAccessCardUserInfoCourse.EmplogeeID = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfoCourse.CHNName = objReader.GetValue(2)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfoCourse.Name = objReader.GetValue(1)
                    'If Not objReader.GetValue(3) Is DBNull.Value Then GetAccessCardUserInfo.Dept = objReader.GetValue(3)
                End While
            Else
                While objReader.Read()
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetAccessCardUserInfoCourse.AccessCardID = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetAccessCardUserInfoCourse.EmplogeeID = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfoCourse.CHNName = objReader.GetValue(2)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetAccessCardUserInfoCourse.Name = objReader.GetValue(1)
                End While
            End If
        Catch ex As Exception
            ErrorLogging("TDC-GetAccessCardUserInfoCourse", "", AccessCardID & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    Public Function GetMGTraceData(ByVal MaterialGroup As String) As MGTraceData

        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        GetMGTraceData.TraceabilityLevel = "NT"     '   Default is Non traceable
        GetMGTraceData.AddlData = ""

        CLMasterSQLCommand = New SqlClient.SqlCommand("Select TraceabilityLevel, AddlData from T_MaterialGroups where MaterialGroup=@MaterialGroup", myConn)
        CLMasterSQLCommand.Parameters.Add("@MaterialGroup", SqlDbType.VarChar, 10, "MaterialGroup")
        CLMasterSQLCommand.Parameters("@MaterialGroup").Value = MaterialGroup

        Try
            myConn.Open()
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetMGTraceData.TraceabilityLevel = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetMGTraceData.AddlData = objReader.GetValue(1)
            End While
            objReader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-GetMGTraceData", "", "MaterialGroup: " & MaterialGroup & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function GetMGTraceLevel(ByVal MaterialNo As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetMGTraceLevel  '{0}' ", MaterialNo)
                GetMGTraceLevel = Convert.ToString(da.ExecuteScalar(Sqlstr))
                Return GetMGTraceLevel

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetMGTraceLevel", "", "MaterialNo: " & MaterialNo & ", " & ex.Message & ex.Source, "E")
                GetMGTraceLevel = "NT"
            End Try
        End Using

    End Function

    'Public Function GetAML(ByVal ItemList() As String) As DataSet
    '    Using da As DataAccess = GetDataAccess()

    '        Dim myAML As DataSet = New DataSet

    '        Dim AMLData As DataTable
    '        Dim ItemData As DataTable
    '        Dim myDR As Data.DataRow

    '        AMLData = New Data.DataTable("AMLData")
    '        AMLData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
    '        AMLData.Columns.Add(New Data.DataColumn("MFR", System.Type.GetType("System.String")))
    '        AMLData.Columns.Add(New Data.DataColumn("MPN", System.Type.GetType("System.String")))
    '        AMLData.Columns.Add(New Data.DataColumn("AMLStatus", System.Type.GetType("System.String")))
    '        AMLData.Columns.Add(New Data.DataColumn("Match", System.Type.GetType("System.String")))
    '        myAML.Tables.Add(AMLData)

    '        ItemData = New Data.DataTable("ItemData")
    '        ItemData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("Stemp", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("MSL", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("COC", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("RoHS", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("AddlData", System.Type.GetType("System.String")))
    '        ItemData.Columns.Add(New Data.DataColumn("EER", System.Type.GetType("System.String")))
    '        myAML.Tables.Add(ItemData)

    '        Dim ds As DataSet = New DataSet

    '        Try
    '            Dim Sqlstr, AMLFlag As String

    '            Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'IPROAML'")
    '            AMLFlag = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
    '            If AMLFlag = "0" Then
    '                Try
    '                    ds = GetAMLLocal(ItemList, myAML)
    '                Catch ex As Exception
    '                    ErrorLogging("GetAML-Call: GetAMLLocal", "", ex.Message & ex.Source, "E")
    '                    Return myAML
    '                End Try
    '                Return ds
    '            End If


    '            'Read AML Data from iPro Production Server
    '            Dim iProWS As iProWS.FusionPOC = New iProWS.FusionPOC
    '            Dim AMLList As iProWS.VeTraceAmlBase()

    '            iProWS.Credentials = System.Net.CredentialCache.DefaultCredentials
    '            iProWS.Timeout = 1000 * 60 * 30

    '            AMLList = iProWS.GetAMLByPartList(ItemList)

    '            If AMLList Is Nothing OrElse AMLList.Length = 0 Then
    '            Else
    '                Dim i As Integer
    '                For i = 0 To AMLList.Length - 1
    '                    Dim MaterialNo As String = AMLList(i).PartNumber.ToString

    '                    myDR = AMLData.NewRow()
    '                    myDR("MaterialNo") = MaterialNo
    '                    myDR("MFR") = FixNull(AMLList(i).Name).ToUpper                  'Manufacturer.ToString            
    '                    myDR("MPN") = FixNull(AMLList(i).ManufacturerPart).ToUpper
    '                    myDR("AMLStatus") = FixNull(AMLList(i).AmlStatus).ToUpper
    '                    myDR("Match") = DBNull.Value
    '                    AMLData.Rows.Add(myDR)

    '                    Dim DR() As DataRow = Nothing
    '                    DR = ItemData.Select(" MaterialNo = '" & MaterialNo & "'")
    '                    If DR.Length = 0 Then
    '                        myDR = ItemData.NewRow()
    '                        myDR("MaterialNo") = MaterialNo
    '                        myDR("Stemp") = FixNull(AMLList(i).Temperature).ToUpper
    '                        myDR("MSL") = FixNull(AMLList(i).MoistureSensitivityLevel).ToUpper
    '                        'myDR("COC") = FixNull(AMLList(i).CertificateOfCompliance).ToUpper
    '                        myDR("RoHS") = FixNull(AMLList(i).RoHs).ToUpper
    '                        'myDR("EER") = FixNull(AMLList(i).EERReference).ToUpper                       'Read EER from iPro here   --10/12/2016
    '                        myDR("AddlData") = DBNull.Value

    '                        'Convert COCRequired to Yes / No instead of True / False, same as iPro screen displays  -- 09/01/2014
    '                        Dim COC As String = FixNull(AMLList(i).CertificateOfCompliance).ToUpper.Trim
    '                        If COC = "TRUE" Then
    '                            myDR("COC") = "YES"
    '                        Else
    '                            myDR("COC") = "NO"
    '                        End If

    '                        Dim ESD As String = ""
    '                        Dim Safety As String = ""
    '                        Dim AddlData As String = ""
    '                        AddlData = FixNull(AMLList(i).Esd).ToUpper
    '                        If AddlData = "TRUE" Then ESD = "E"

    '                        AddlData = FixNull(AMLList(i).SafetyCritical).ToUpper
    '                        If AddlData = "2" Then Safety = "S"

    '                        AddlData = Safety & ESD
    '                        If AddlData.ToString.Trim <> "" Then myDR("AddlData") = AddlData

    '                        ItemData.Rows.Add(myDR)
    '                    End If
    '                Next
    '            End If

    '            Return myAML

    '        Catch ex As Exception
    '            ErrorLogging("PublicFunction-GetAML", "", ex.Message & ex.Source, "E")
    '            Return myAML
    '        End Try
    '    End Using

    'End Function
    Public Function GetAML(ByVal ItemList() As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim myAML As DataSet = New DataSet

            Dim AMLData As DataTable
            Dim ItemData As DataTable
            Dim myDR As Data.DataRow

            AMLData = New Data.DataTable("AMLData")
            AMLData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
            AMLData.Columns.Add(New Data.DataColumn("MFR", System.Type.GetType("System.String")))
            AMLData.Columns.Add(New Data.DataColumn("MPN", System.Type.GetType("System.String")))
            AMLData.Columns.Add(New Data.DataColumn("AMLStatus", System.Type.GetType("System.String")))
            AMLData.Columns.Add(New Data.DataColumn("Match", System.Type.GetType("System.String")))
            AMLData.Columns.Add(New Data.DataColumn("EER", System.Type.GetType("System.String")))                 'EER is related to MFR/MPN
            myAML.Tables.Add(AMLData)

            ItemData = New Data.DataTable("ItemData")
            ItemData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("Stemp", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("MSL", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("COC", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("RoHS", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("AddlData", System.Type.GetType("System.String")))
            'ItemData.Columns.Add(New Data.DataColumn("EER", System.Type.GetType("System.String")))
            ItemData.Columns.Add(New Data.DataColumn("RESData", System.Type.GetType("System.String")))                      'Revision & ESD & ShelfLife

            myAML.Tables.Add(ItemData)

            Dim i As Integer
            Dim dtAML As New DataTable

            Try
                Dim Sqlstr, AMLFlag As String
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'IPROAML'")
                AMLFlag = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper

                If AMLFlag = "0" Then
                    Dim PartNo As String = ""
                    Dim MatlList As String = ""

                    For i = 0 To ItemList.Length - 1
                        If ItemList(i) Is Nothing OrElse ItemList(i).ToString = "" Then
                        ElseIf ItemList(i).ToString <> "" Then
                            If PartNo = "" Then
                                PartNo = "''" & ItemList(i).ToString & "''"
                            Else
                                PartNo = PartNo & ",''" & ItemList(i).ToString & "''"
                            End If
                        End If
                    Next
                    MatlList = "(" & PartNo & ")"

                    Try
                        Sqlstr = String.Format("exec sp_GetAML '{0}'", MatlList)
                        dtAML = da.ExecuteDataTable(Sqlstr, 1)    '0 = eTrace DataBase, 1 = Local eTraceAML Database
                    Catch ex As Exception
                        ErrorLogging("GetAML-Call: sp_GetAML", "", ex.Message & ex.Source, "E")
                        Return myAML
                    End Try

                Else

                    'Read AML Data from iPro Production Server
                    Dim iProWS As iProWS.FusionPOC = New iProWS.FusionPOC
                    'Dim AMLList As iProWS.VeTraceAmlBase()

                    iProWS.Credentials = System.Net.CredentialCache.DefaultCredentials
                    iProWS.Timeout = 1000 * 60 * 30

                    Dim ds As DataSet = New DataSet
                    Try
                        ds = iProWS.GetAMLByPartList(ItemList)
                    Catch ex As Exception
                        ErrorLogging("PublicFunction-GetAML1", "", ex.Message & ex.Source, "E")
                        Return Nothing
                    End Try
                    dtAML = ds.Tables(0).Copy
                End If


                If dtAML Is Nothing OrElse dtAML.Rows.Count = 0 Then
                ElseIf dtAML.Rows.Count > 0 Then
                    'Dim i As Integer
                    For i = 0 To dtAML.Rows.Count - 1
                        Dim MaterialNo As String = dtAML.Rows(i)("PartNumber").ToString

                        myDR = myAML.Tables("AMLData").NewRow()
                        myDR("MaterialNo") = MaterialNo
                        myDR("MFR") = FixNull(dtAML.Rows(i)("Name")).ToUpper
                        myDR("MPN") = FixNull(dtAML.Rows(i)("ManufacturerPart")).ToUpper
                        myDR("AMLStatus") = FixNull(dtAML.Rows(i)("AmlStatus")).ToUpper
                        myDR("Match") = DBNull.Value
                        myDR("EER") = FixNull(dtAML.Rows(i)("EERReference")).ToUpper                              'Read EER from iPro here   --10/26/2016
                        myAML.Tables("AMLData").Rows.Add(myDR)

                        Dim DR() As DataRow = Nothing
                        DR = myAML.Tables("ItemData").Select(" MaterialNo = '" & MaterialNo & "'")
                        If DR.Length = 0 Then
                            myDR = myAML.Tables("ItemData").NewRow()
                            myDR("MaterialNo") = MaterialNo
                            myDR("Stemp") = FixNull(dtAML.Rows(i)("Temperature")).ToUpper
                            myDR("MSL") = FixNull(dtAML.Rows(i)("MoistureSensitivityLevel")).ToUpper
                            'myDR("COC") = FixNull(dtAML.Rows(i)("CertificateOfCompliance")).ToUpper
                            myDR("RoHS") = FixNull(dtAML.Rows(i)("RoHs")).ToUpper
                            'myDR("EER") = FixNull(dtAML.Rows(i)("EERReference")).ToUpper                              'Read EER from iPro here   --10/12/2016
                            myDR("AddlData") = DBNull.Value
                            myDR("RESData") = ""

                            'Convert COCRequired to Yes / No instead of True / False, same as iPro screen displays  -- 09/01/2014
                            Dim COC As String = FixNull(dtAML.Rows(i)("CertificateOfCompliance")).ToUpper.Trim
                            If COC = "TRUE" Then
                                myDR("COC") = "YES"
                            Else
                                myDR("COC") = "NO"
                            End If

                            Dim ESD As String = ""
                            Dim Safety As String = ""
                            Dim AddlData As String = ""
                            AddlData = FixNull(dtAML.Rows(i)("Esd")).ToUpper
                            If AddlData = "TRUE" Then ESD = "E"
                            myDR("RESData") = FixNull(dtAML.Rows(i)("PartRevision")).ToUpper & "~" & ESD & "~" & FixNull(dtAML.Rows(i)("ShelfLife"))   '05/04/2017

                            AddlData = FixNull(dtAML.Rows(i)("SafetyCritical")).ToUpper
                            If AddlData = "2" Then Safety = "S"

                            AddlData = Safety & ESD
                            If AddlData.ToString.Trim <> "" Then myDR("AddlData") = AddlData

                            myAML.Tables("ItemData").Rows.Add(myDR)
                        End If
                    Next
                End If

                Return myAML

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetAML", "", ex.Message & ex.Source, "E")
                Return myAML
            End Try
        End Using

    End Function

    Public Function GetAMLLocal(ByVal ItemList() As String, ByVal myAML As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim PartNo As String = ""
            Dim MatlList As String = ""
            Dim myDR As Data.DataRow

            Dim i As Integer
            Dim Sqlstr As String
            Dim dtAML As New DataTable

            Try

                For i = 0 To ItemList.Length - 1
                    If ItemList(i) Is Nothing OrElse ItemList(i).ToString = "" Then
                    ElseIf ItemList(i).ToString <> "" Then
                        If PartNo = "" Then
                            PartNo = "''" & ItemList(i).ToString & "''"
                        Else
                            PartNo = PartNo & ",''" & ItemList(i).ToString & "''"
                        End If
                    End If
                Next
                MatlList = "(" & PartNo & ")"

                Sqlstr = String.Format("exec sp_GetAML '{0}'", MatlList)
                dtAML = da.ExecuteDataTable(Sqlstr, 1)    '0 = eTrace DataBase, 1 = Local eTraceAML Database

                If dtAML Is Nothing OrElse dtAML.Rows.Count = 0 Then
                ElseIf dtAML.Rows.Count > 0 Then
                    For i = 0 To dtAML.Rows.Count - 1
                        Dim MaterialNo As String = dtAML.Rows(i)("PartNumber").ToString

                        myDR = myAML.Tables("AMLData").NewRow()
                        myDR("MaterialNo") = MaterialNo
                        myDR("MFR") = FixNull(dtAML.Rows(i)("Name")).ToUpper
                        myDR("MPN") = FixNull(dtAML.Rows(i)("ManufacturerPart")).ToUpper
                        myDR("AMLStatus") = FixNull(dtAML.Rows(i)("AmlStatus")).ToUpper
                        myDR("Match") = DBNull.Value
                        myDR("EER") = FixNull(dtAML.Rows(i)("EERReference")).ToUpper                              'Read EER from iPro here   --10/26/2016
                        myAML.Tables("AMLData").Rows.Add(myDR)

                        Dim DR() As DataRow = Nothing
                        DR = myAML.Tables("ItemData").Select(" MaterialNo = '" & MaterialNo & "'")
                        If DR.Length = 0 Then
                            myDR = myAML.Tables("ItemData").NewRow()
                            myDR("MaterialNo") = MaterialNo
                            myDR("Stemp") = FixNull(dtAML.Rows(i)("Temperature")).ToUpper
                            myDR("MSL") = FixNull(dtAML.Rows(i)("MoistureSensitivityLevel")).ToUpper
                            'myDR("COC") = FixNull(dtAML.Rows(i)("CertificateOfCompliance")).ToUpper
                            myDR("RoHS") = FixNull(dtAML.Rows(i)("RoHs")).ToUpper
                            'myDR("EER") = FixNull(dtAML.Rows(i)("EERReference")).ToUpper                              'Read EER from iPro here   --10/12/2016
                            myDR("AddlData") = DBNull.Value
                            myDR("RESData") = ""

                            'Convert COCRequired to Yes / No instead of True / False, same as iPro screen displays  -- 09/01/2014
                            Dim COC As String = FixNull(dtAML.Rows(i)("CertificateOfCompliance")).ToUpper.Trim
                            If COC = "TRUE" Then
                                myDR("COC") = "YES"
                            Else
                                myDR("COC") = "NO"
                            End If

                            Dim ESD As String = ""
                            Dim Safety As String = ""
                            Dim AddlData As String = ""
                            AddlData = FixNull(dtAML.Rows(i)("Esd")).ToUpper
                            If AddlData = "TRUE" Then ESD = "E"
                            myDR("RESData") = FixNull(dtAML.Rows(i)("PartRevision")).ToUpper & "~" & ESD & "~" & FixNull(dtAML.Rows(i)("ShelfLife"))   '05/04/2017

                            AddlData = FixNull(dtAML.Rows(i)("SafetyCritical")).ToUpper
                            If AddlData = "2" Then Safety = "S"

                            AddlData = Safety & ESD
                            If AddlData.ToString.Trim <> "" Then myDR("AddlData") = AddlData

                            myAML.Tables("ItemData").Rows.Add(myDR)
                        End If
                    Next
                End If

                Return myAML

            Catch ex As Exception
                ErrorLogging("PublicFunction-GetAMLLocal", "", ex.Message & ex.Source, "E")
                Return myAML
            End Try

        End Using

    End Function

    Public Function GetAMLData(ByVal ItemList() As String, ByVal ModuleName As String) As DataSet

        Dim myAML As DataSet = Nothing

        Dim AMLStatus As String = ""
        Dim DR() As DataRow = Nothing

        Try

            Dim k As Integer = 0

            'Read AML Data from iPro 
            While (k < 3 And myAML Is Nothing)
                Try
                    myAML = New DataSet
                    myAML = GetAML(ItemList)
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("PublicFunction-GetAMLData" & Str(k), Str(k), "Module Name: " & ModuleName & "; " & ex.Message & ex.Source, "E")
                    myAML = Nothing
                End Try
            End While


            If myAML Is Nothing OrElse myAML.Tables.Count = 0 OrElse myAML.Tables(0).Rows.Count = 0 Then
                Return myAML
            End If

            Dim Sqlstr As String
            Dim da As DataAccess = GetDataAccess()
            If ModuleName.ToUpper = "RECEIVING" Then
                'AMLStatus = "ACTIVE"                       'AMLStatus will seprate by Comma
                'Which AML Status can be extracted from iPro for Receiving ?
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC003'")
                AMLStatus = Convert.ToString(da.ExecuteScalar(Sqlstr))

            ElseIf ModuleName.ToUpper = "ITEM TRANSFER" Then
                'Which AML Status can be extracted from iPro for Item Transfer ?    Only two status: ACTIVE,SUPPLIER DISCONTINUED
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID012'")
                AMLStatus = Convert.ToString(da.ExecuteScalar(Sqlstr))
            End If

            If AMLStatus <> "" Then
                DR = myAML.Tables("AMLData").Select("AMLStatus NOT IN ( '" & AMLStatus.Replace(",", "','") & "') ")
                If DR.Length > 0 Then
                    Dim i As Integer
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    myAML.Tables("AMLData").AcceptChanges()
                End If
            End If

            Return myAML

        Catch ex As Exception
            ErrorLogging("PublicFunction-GetAMLData", "", "Module Name: " & ModuleName & "; " & ex.Message & ex.Source, "E")
            Return myAML
        End Try

    End Function

    Public Function Check_OSPUser(ByVal LoginData As ERPLogin) As UserData
        Using da As DataAccess = GetDataAccess()

            Check_OSPUser = New UserData

            Try
                Dim a, b As Integer
                Dim OraStr, OraServer As String
                OraStr = da._OConnString
                a = InStr(13, OraStr, ";", CompareMethod.Text)
                If a > 0 Then
                    b = a - 13
                    OraServer = Mid(OraStr, 13, b)
                Else
                    OraServer = "Oracle"
                End If

                Dim Sqlstr As String
                Dim myUsers = New DataSet

                Sqlstr = String.Format("Select UserID, Name, EmpID, Dept, UserType from T_Users with (nolock) where Status = 'Active' and UserID = '{0}'", LoginData.User)
                myUsers = da.ExecuteDataSet(Sqlstr, "Users")
                If myUsers Is Nothing OrElse myUsers.Tables.Count = 0 OrElse myUsers.Tables(0).Rows.Count = 0 Then
                    Check_OSPUser.ErrorMsg = "Invalid User"
                    Exit Function
                End If

                Dim UserType As String = myUsers.Tables(0).Rows(0)("UserType").ToString.ToUpper
                If UserType <> "OSP" Then Exit Function


                Dim UserID As String = ""
                Sqlstr = String.Format("Select UserID from T_UserPWD with (nolock) where (ValidFrom <= CONVERT(DATETIME, getdate(), 102)) and (ValidTo >= CONVERT(DATETIME, getdate(), 102)) and UserID = '{0}' and Password = '{1}'", LoginData.User, LoginData.PWD)
                UserID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                If UserID = "" Then
                    Check_OSPUser.Server = OraServer
                    Check_OSPUser.User = LoginData.User.ToUpper
                    Check_OSPUser.ErrorMsg = "Login failed, check OSP User/pwd"
                    Exit Function
                End If

                Check_OSPUser.Server = OraServer
                Check_OSPUser.User = UserID.ToUpper
                Check_OSPUser.UserID = myUsers.Tables(0).Rows(0)("EmpID").ToString
                Check_OSPUser.UserDept = myUsers.Tables(0).Rows(0)("Dept").ToString.ToUpper
                Check_OSPUser.UserType = myUsers.Tables(0).Rows(0)("UserType").ToString.ToUpper
                Check_OSPUser.FirstName = myUsers.Tables(0).Rows(0)("Name").ToString.ToUpper
                Check_OSPUser.LastName = ""
                Check_OSPUser.ErrorMsg = ""

                Check_OSPUser.RespID_PO = "OSP"
                Check_OSPUser.RespID_Inv = "OSP"
                Check_OSPUser.RespID_WIP = "OSP"

            Catch ex As Exception
                ErrorLogging("PublicFunction-Check_OSP_User", "", "User: " & LoginData.User & ex.Message & ex.Source, "E")
                Check_OSPUser.ErrorMsg = "Invalid User"
            End Try
        End Using

    End Function

    'Public Function CleanBoxID(ByVal CLID As String, ByVal ERPLoginData As ERPLogin) As Boolean
    '    Using da As DataAccess = GetDataAccess()
    '        Try
    '            Dim BoxID As String
    '            Dim sqlstr As String
    '            sqlstr = String.Format("select BoxID from T_CLMaster where CLID = '{0}'", CLID)
    '            BoxID = FixNull(Convert.ToString(da.ExecuteScalar(sqlstr)))
    '            If BoxID <> "" Then
    '                da.ExecuteNonQuery(String.Format("update T_CLMaster set BoxID='' where BoxID='{0}'", BoxID))
    '                Return True
    '            End If
    '        Catch ex As Exception
    '            ErrorLogging("CleanBoxID", ERPLoginData.User, ex.Message & ex.Source, "E")
    '            Return False
    '        End Try
    '    End Using
    'End Function

    Public Function ClearBoxID(ByVal BoxID As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                If BoxID <> "" Then
                    da.ExecuteNonQuery(String.Format("update T_CLMaster set BoxID='' where BoxID='{0}'", BoxID))
                    Return True
                End If
            Catch ex As Exception
                ErrorLogging("ClearBoxID", ERPLoginData.User, ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function GetSysMessage(ByVal frequecy As String, ByVal device As String, ByVal eTraceModel As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String = String.Empty

                If (String.Equals(device, "PC", StringComparison.OrdinalIgnoreCase)) Then
                    If (Not String.IsNullOrEmpty(eTraceModel)) Then
                        If (String.Equals(eTraceModel, "TestDataCollection", StringComparison.OrdinalIgnoreCase) Or _
                            String.Equals(eTraceModel, "ShopFloorControl", StringComparison.OrdinalIgnoreCase)) Then
                            sqlstr = String.Format(String.Format("Select * FROM T_SysMsg  with (nolock) WHERE MsgFreq='{0}' AND PC_SF='{1}' AND ValidFrom<=getdate() AND Expiry>getdate() ", frequecy, "True"))

                        Else
                            sqlstr = String.Format(String.Format("Select * FROM T_SysMsg  with (nolock) WHERE MsgFreq='{0}' AND PC_MC='{1}' AND ValidFrom<=getdate() AND Expiry>getdate() ", frequecy, "True"))
                        End If
                    End If

                ElseIf (String.Equals(device, "HH", StringComparison.OrdinalIgnoreCase)) Then
                    sqlstr = String.Format(String.Format("Select * FROM T_SysMsg  with (nolock) WHERE MsgFreq='{0}' AND HH='{1}' AND ValidFrom<=getdate() AND Expiry>getdate() ", frequecy, "True"))
                End If

                Return da.ExecuteDataSet(sqlstr)
            Catch ex As Exception
                ErrorLogging("GetSysMessage", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function ShowMessage(ByVal MessageText As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim OriginalText As String
                Dim Target1 As String = "^"
                Dim Target2 As String = "@"
                Dim between As String
                Dim before As String
                Dim pos As Long
                OriginalText = MessageText               '      ^REC-0@"

                ' Get the text before target1.
                pos = InStr(MessageText, Target1)
                If pos = 0 Then
                    ' target1 is missing. Set before = "".
                    before = ""
                Else
                    ' Set before.
                    before = Microsoft.VisualBasic.Left$(MessageText, pos - 1)

                    ' Remove up to target1 from the string.
                    MessageText = Mid$(MessageText, pos + Len(Target1))
                End If

                ' Get the text before target2.
                pos = InStr(MessageText, Target2)
                If pos = 0 Then
                    ' target2 is missing. Set between = "".
                    between = ""
                Else
                    ' Set between.
                    between = Microsoft.VisualBasic.Left$(MessageText, pos - 1)
                    ' Remove up to target2 from the string.
                    MessageText = Mid$(MessageText, pos + Len(Target2))
                End If

                If between <> "" Then
                    Dim Lang, SqlStr As String
                    SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'SYS002' ")
                    Lang = Convert.ToString(da.ExecuteScalar(SqlStr)).ToUpper

                    Dim Message As String = ""
                    Dim MessageClass As String = Microsoft.VisualBasic.Left(between, 3)
                    Dim MessageID As Integer = Microsoft.VisualBasic.Mid(between, 5)

                    SqlStr = String.Format("Select Message from T_Messages with (nolock) where MessageClass = '{0}' and MessageID = '{1}' and Lang = '{2}' ", MessageClass, MessageID, Lang)
                    Message = Convert.ToString(da.ExecuteScalar(SqlStr)).ToUpper
                    If Message = "" Then
                        ShowMessage = "Message " & MessageText & " is not defined yet. "
                        Exit Function
                    End If

                    OriginalText = OriginalText.Replace("^" & between & "@", Message)
                    OriginalText = ShowMessage(OriginalText)
                End If
                ShowMessage = OriginalText

            Catch ex As Exception
                ErrorLogging("PublicFunction-ShowMessage", "", "Invalid Message: " & MessageText & ", " & ex.Message & ex.Source, "E")
                ShowMessage = "Invalid Message " & MessageText
            End Try
        End Using

    End Function

    Public Function ValidatePallet(ByVal LoginData As ERPLogin, ByVal PalletID As String, ByVal MatlStatus As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim Status As String = ""

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select TOP (1) Status from T_RecDashboard with (nolock) where PalletID = '{0}' AND Status <> '99' ", PalletID)
                Status = Convert.ToString(da.ExecuteScalar(Sqlstr))
                If Status = "" Then Return "Y"

                Status = ""
                If MatlStatus = "RECEIVING" Then
                    Sqlstr = String.Format("Select TOP (1) Status from T_RecDashboard with (nolock) where PalletID = '{0}' AND Status <> '10' ", PalletID)
                    Status = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If Status = "" Then Return "Y"
                Else
                    Sqlstr = String.Format("Select TOP (1) Status from T_RecDashboard with (nolock) where PalletID = '{0}' AND ( Status <> '99' and Status <> '{1}' ) ", PalletID, MatlStatus)
                    Status = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If Status = "" Then Return "Y"
                End If

                Dim StaDesc As String
                Sqlstr = String.Format("Select Description from T_SysLOV with (nolock) where Name='Receiving Dashboard' AND ProcessID = '{0}' ", Status)
                StaDesc = Convert.ToString(da.ExecuteScalar(Sqlstr))
                ValidatePallet = "The PalletID is currently in use, and now its status is: " & StaDesc

                If MatlStatus = "RECEIVING" Then
                    ValidatePallet = "^REC-99@" & " " & StaDesc
                End If

            Catch ex As Exception
                ErrorLogging("PublicFunction-ValidatePallet", LoginData.User, "Validate error for PalletID " & PalletID & ", " & ex.Message & ex.Source, "E")
                ValidatePallet = "Validate error for PalletID " & PalletID
            End Try
        End Using

    End Function

    Public Function ValidateBerth(ByVal LoginData As ERPLogin, ByVal BerthID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim PalletID As String = ""

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select TOP (1) PalletID from T_RecDashboard with (nolock) where BerthID = '{0}' AND Status <> '99' ", BerthID)
                PalletID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                If PalletID = "" Then Return "Y"

                ValidateBerth = "The BerthID is currently in use by Pallet " & PalletID

            Catch ex As Exception
                ErrorLogging("PublicFunction-ValidateBerth", LoginData.User, "Validate error for BerthID " & BerthID & ", " & ex.Message & ex.Source, "E")
                ValidateBerth = "Validate error for PalletID " & BerthID
            End Try
        End Using

    End Function

    Public Function PrintInterOrgCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            PrintInterOrgCLIDs = True
            If CLIDs Is Nothing OrElse CLIDs.Tables.Count = 0 OrElse CLIDs.Tables(0).Rows.Count = 0 Then
                PrintInterOrgCLIDs = False
                Exit Function
            End If

            Dim TmpDate As Date
            Dim lblData As MatLabel = New MatLabel
            Dim PCBLabel As POLabel = New POLabel

            Dim i As Integer
            Dim LabelType, lblPrint, StartStr, MidStr As String
            Dim CLID, DestOrg, User, LotCtrlCode, Sqlstr As String

            User = CLIDs.Tables(0).Rows(0)("User").ToString
            DestOrg = CLIDs.Tables(0).Rows(0)("DestOrg").ToString
            LotCtrlCode = CLIDs.Tables(0).Rows(0)("LotCtrlCode").ToString

            Try

                For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                    CLID = CLIDs.Tables(0).Rows(i)("CLID").ToString

                    Dim LblStart As String = Microsoft.VisualBasic.Left(CLID, 2)         'LD Huawei requirement -- 11/16/2016

                    StartStr = Microsoft.VisualBasic.Left(CLID, 1)
                    MidStr = Microsoft.VisualBasic.Mid(CLID, 3, 1)

                    If StartStr = "B" OrElse MidStr = "P" OrElse LblStart = "LE" Then
                        LabelType = "PCBID"
                        Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo,MaterialRevision,PurOrdNo,QtyBaseUOM,BaseUOM,RecDate,StatusCode,SONo,SOLine, LotNo, DateCode, ManufacturerPN, RoHS from T_CLMaster with (nolock) where CLID = '{0}' ", CLID)
                        'ElseIf MidStr = "B" Then
                        '    Sqlstr = String.Format("Select BoxID,OrgCode,MaterialNo,MaterialRevision,Sum(QtyBaseUOM) As BoxQty,BaseUOM,'' as LotNo,'' as DateCode,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot from T_CLMaster group by BoxID,OrgCode,MaterialNo,MaterialRevision,BaseUOM,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot having StatusCode > 0 AND BoxID = '{0}' ", CLID)
                        '    lblData.LabelType = "Box Label"
                        '    LabelType = "BoxID"
                    ElseIf IsNumeric(MidStr) = True OrElse MidStr = "V" Then
                        Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo,MaterialRevision,QtyBaseUOM,BaseUOM,LotNo,DateCode,ExpDate,RecDate,RecDocNo,RoHS,StockType,MatSuffix1,MatSuffix2,MatSuffix3,StatusCode,AddlData,Manufacturer,ManufacturerPN,Stemp,MSL,RTLot,MCPosition, PredefinedSubInv, ItemText, CountryOfOrigin as COO from T_CLMaster with (nolock) where CLID = '{0}' ", CLID)
                        lblData.LabelType = "Material Label"
                        LabelType = "CLID"
                    Else
                        ErrorLogging("PublicFunction-PrintInterOrgCLIDs", "", "CLID: " & CLID & " not in the label type list ", "I")
                        PrintInterOrgCLIDs = False
                        Continue For
                    End If

                    Dim ds As DataSet = New DataSet
                    ds = da.ExecuteDataSet(Sqlstr, "CLIDs")

                    If ds.Tables(0).Rows.Count = 0 Then
                        ErrorLogging("PublicFunction-PrintInterOrgCLIDs", "", "CLID: " & CLID & " not found ", "I")
                        PrintInterOrgCLIDs = False
                        Continue For
                    End If

                    If LabelType = "PCBID" Then
                        ' Record PCB Label Info
                        PCBLabel.OrgCode = DestOrg
                        If Not ds.Tables(0).Rows(0)(0) Is DBNull.Value Then PCBLabel.CLID = ds.Tables(0).Rows(0)(0)
                        'If Not ds.Tables(0).Rows(0)(1) Is DBNull.Value Then PCBLabel.OrgCode = ds.Tables(0).Rows(0)(1) 'OrgCode
                        If Not ds.Tables(0).Rows(0)(2) Is DBNull.Value Then PCBLabel.PCBA = ds.Tables(0).Rows(0)(2)
                        If Not ds.Tables(0).Rows(0)(3) Is DBNull.Value Then PCBLabel.Rev = ds.Tables(0).Rows(0)(3)
                        If Not ds.Tables(0).Rows(0)(4) Is DBNull.Value Then PCBLabel.POID = ds.Tables(0).Rows(0)(4)

                        Dim POQty As Integer = 0
                        If Not ds.Tables(0).Rows(0)(5) Is DBNull.Value Then POQty = CInt(ds.Tables(0).Rows(0)(5))
                        PCBLabel.Qty = POQty.ToString

                        If Not ds.Tables(0).Rows(0)(6) Is DBNull.Value Then PCBLabel.UOM = ds.Tables(0).Rows(0)(6)
                        If Not ds.Tables(0).Rows(0)(7) Is DBNull.Value Then
                            TmpDate = ds.Tables(0).Rows(0)(7)
                            PCBLabel.RecDate = TmpDate.ToString("MM/dd/yyyy")
                        End If

                        If Not ds.Tables(0).Rows(0)(9) Is DBNull.Value Then PCBLabel.SONo = ds.Tables(0).Rows(0)(9)
                        If Not ds.Tables(0).Rows(0)(10) Is DBNull.Value Then PCBLabel.SOLine = ds.Tables(0).Rows(0)(10)
                        If Not ds.Tables(0).Rows(0)(11) Is DBNull.Value Then PCBLabel.LotNo = ds.Tables(0).Rows(0)(11) ''Add Lot No to PCB Lable.

                        '' Add DateCode to PCBA Lable.
                        If Not ds.Tables(0).Rows(0)(12) Is DBNull.Value Then PCBLabel.DateCode = ds.Tables(0).Rows(0)(12)

                        'Add MPN to PCBA Label for Mix Source Build Project   09/09/2013            
                        If Not ds.Tables(0).Rows(0)(13) Is DBNull.Value Then PCBLabel.MPN = ds.Tables(0).Rows(0)(13)

                        'Add RoHS to PCBA Label    05/19/2015            
                        If Not ds.Tables(0).Rows(0)(14) Is DBNull.Value Then PCBLabel.RoHS = ds.Tables(0).Rows(0)(14)

                        lblPrint = PrintPOLabel(Printer, PCBLabel)

                    Else

                        lblData.OrgCode = DestOrg
                        If Not ds.Tables(0).Rows(0)(0) Is DBNull.Value Then lblData.LabelID = ds.Tables(0).Rows(0)(0)
                        'If Not ds.Tables(0).Rows(0)(1) Is DBNull.Value Then lblData.OrgCode = ds.Tables(0).Rows(0)(1)
                        If Not ds.Tables(0).Rows(0)(2) Is DBNull.Value Then lblData.Material = ds.Tables(0).Rows(0)(2)
                        If Not ds.Tables(0).Rows(0)(3) Is DBNull.Value Then lblData.Rev = ds.Tables(0).Rows(0)(3)
                        If Not ds.Tables(0).Rows(0)(4) Is DBNull.Value Then lblData.Qty = ds.Tables(0).Rows(0)(4)
                        If Not ds.Tables(0).Rows(0)(5) Is DBNull.Value Then lblData.UOM = ds.Tables(0).Rows(0)(5)
                        If Not ds.Tables(0).Rows(0)(6) Is DBNull.Value Then lblData.LotNo = ds.Tables(0).Rows(0)(6)
                        If Not ds.Tables(0).Rows(0)(7) Is DBNull.Value Then lblData.DCode = ds.Tables(0).Rows(0)(7)

                        lblData.Qty = Format(ds.Tables(0).Rows(0)(4), "#0.#####")

                        If Not ds.Tables(0).Rows(0)(8) Is DBNull.Value Then
                            TmpDate = ds.Tables(0).Rows(0)(8)
                            lblData.ExpDate = TmpDate.ToString("MM/dd/yyyy")

                            'No LotConrol, should keep ExpDate / RTLot as blank
                            If LotCtrlCode = "1" Then lblData.ExpDate = ""
                        End If

                        If Not ds.Tables(0).Rows(0)(9) Is DBNull.Value Then
                            TmpDate = ds.Tables(0).Rows(0)(9)
                            lblData.RecDate = TmpDate.ToString("MM/dd/yyyy")
                        End If

                        'Default set to print out Barcode for Material
                        lblData.IPN = lblData.Material

                        If Not ds.Tables(0).Rows(0)(10) Is DBNull.Value Then lblData.RT = ds.Tables(0).Rows(0)(10)
                        If Not ds.Tables(0).Rows(0)(11) Is DBNull.Value Then lblData.RoHS = ds.Tables(0).Rows(0)(11)
                        If Not ds.Tables(0).Rows(0)(12) Is DBNull.Value Then lblData.InspFlag = ds.Tables(0).Rows(0)(12)
                        If Not ds.Tables(0).Rows(0)(18) Is DBNull.Value Then lblData.Mfr = ds.Tables(0).Rows(0)(18)
                        If Not ds.Tables(0).Rows(0)(19) Is DBNull.Value Then lblData.MfrPN = ds.Tables(0).Rows(0)(19)
                        If Not ds.Tables(0).Rows(0)(20) Is DBNull.Value Then lblData.Stemp = ds.Tables(0).Rows(0)(20)
                        If Not ds.Tables(0).Rows(0)(21) Is DBNull.Value Then lblData.MSL = ds.Tables(0).Rows(0)(21)

                        'Check if RTLot is different from RT, if yes, then Combine RT / RTLot into field lblData.RT
                        If ds.Tables(0).Rows(0)(22) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(0)(22)) = "" Then
                        ElseIf LotCtrlCode <> "1" Then
                            'No LotConrol, should keep ExpDate / RTLot as blank
                            Dim RT As String = lblData.RT
                            Dim RTLot As String = Trim(ds.Tables(0).Rows(0)(22))
                            'If RTLot <> RT Then lblData.RT = lblData.RT & " / " & RTLot
                            If RTLot <> RT Then                                                 'RTLot generated by RT-ExpDate (yyMMdd)
                                If RTLot.Contains("-") Then
                                    lblData.RT = RTLot
                                Else
                                    lblData.RT = lblData.RT & " / " & RTLot
                                End If
                            End If
                        End If

                        'Print MCPosition on MaterialLabel requested by ZS   --- 06/24/2014
                        If Not ds.Tables(0).Rows(0)(23) Is DBNull.Value Then lblData.MCPosition = ds.Tables(0).Rows(0)(23)

                        'Print PredefinedSubInv on MaterialLabel requested by ZS only  --- 03/27/2015
                        If Not ds.Tables(0).Rows(0)(24) Is DBNull.Value Then lblData.PreSubInv = ds.Tables(0).Rows(0)(24)

                        'Print ItemText on MaterialLabel requested by LD Huawei project  --- 12/27/2016
                        If Not ds.Tables(0).Rows(0)(25) Is DBNull.Value Then lblData.ItemText = ds.Tables(0).Rows(0)(25)

                        'Print COO on MaterialLabel requested by ZS only  --- 12/27/2016
                        If Not ds.Tables(0).Rows(0)(26) Is DBNull.Value Then lblData.COO = ds.Tables(0).Rows(0)(26)

                        lblData.Safety = ""
                        lblData.ESD = BlankESD
                        If Not ds.Tables(0).Rows(0)(17) Is DBNull.Value Then
                            Dim AddlData As String = ds.Tables(0).Rows(0)(17).ToString.ToUpper
                            If InStr(AddlData, "S") > 0 Then lblData.Safety = "Safety"
                            If InStr(AddlData, "E") > 0 Then lblData.ESD = ESDImage
                        End If

                        Dim flag As Integer = 0
                        If ds.Tables(0).Rows(0)(13) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(0)(13)) = "" Then
                        Else
                            lblData.MatSuffix = "[" & ds.Tables(0).Rows(0)(13) & "]"
                            lblPrint = PrintLabel(Printer, lblData)
                            flag = 1
                        End If
                        If ds.Tables(0).Rows(0)(14) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(0)(14)) = "" Then
                        Else
                            lblData.MatSuffix = "[" & ds.Tables(0).Rows(0)(14) & "]"
                            lblPrint = PrintLabel(Printer, lblData)
                            flag = 1
                        End If
                        If ds.Tables(0).Rows(0)(15) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(0)(15)) = "" Then
                        Else
                            lblData.MatSuffix = "[" & ds.Tables(0).Rows(0)(15) & "]"
                            lblPrint = PrintLabel(Printer, lblData)
                            flag = 1
                        End If
                        If flag = 0 Then
                            lblPrint = PrintLabel(Printer, lblData)
                        End If

                        If lblPrint <> "True" Then
                            PrintInterOrgCLIDs = False
                        End If
                    End If
                Next

            Catch ex As Exception
                ErrorLogging("PublicFunction-PrintInterOrgCLIDs", User.ToUpper, ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using

    End Function

    Public Function PrinterCheck(ByVal User As String, ByVal PrinterID As String, ByVal OutputType As String) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim LabelType, Sqlstr As String
                Sqlstr = String.Format("exec sp_PrinterCheck '{0}','{1}' ", PrinterID, OutputType)
                Return Convert.ToString(da.ExecuteScalar(Sqlstr))            '.ToUpper()
            Catch ex As Exception
                ErrorLogging("PublicFunction-PrinterCheck", User, "Invalid Printer " & PrinterID & ", " & ex.Message & ex.Source, "E")
                PrinterCheck = "Invalid Printer " & PrinterID
            End Try
        End Using

    End Function

    Public Function PrintCH09Label(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            PrintCH09Label = True

            Dim CH09Flag As Boolean = False
            Dim LabelData As CH09Label = New CH09Label

            Try
                If CLIDs Is Nothing OrElse CLIDs.Tables.Count = 0 OrElse CLIDs.Tables(0).Rows.Count = 0 Then
                    PrintCH09Label = False
                    Exit Function
                End If

                'Check if the dataset table already contained column ItemText, if Yes, then don't need to read data from database table again
                If CLIDs.Tables(0).Columns.Count > 1 AndAlso CLIDs.Tables(0).Columns.Contains("ItemText") Then
                    CH09Flag = True
                End If

                Dim i As Integer
                For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                    Dim CLID As String = CLIDs.Tables(0).Rows(i)(0).ToString

                    If CLID = "" Then
                        ErrorLogging("PublicFunction-PrintCH09Label1", "", "CLID not found ", "I")
                        PrintCH09Label = False
                        Continue For
                    End If

                    Dim Sqlstr As String
                    Dim ds As DataSet = New DataSet

                    If CH09Flag = True Then
                        ds.Merge(CLIDs)
                    Else
                        Sqlstr = String.Format("exec sp_ReadLabelIDs  '{0}','{1}','{2}' ", "", CLID, "CH09Code")
                        ds = da.ExecuteDataSet(Sqlstr, "CLIDs")

                        If ds.Tables(0).Rows.Count = 0 Then
                            ErrorLogging("PublicFunction-PrintCH09Label2", "", "CLID: " & CLID & " not found ", "I")
                            PrintCH09Label = False
                            Continue For
                        End If
                    End If

                    LabelData.CLID = ds.Tables(0).Rows(0)("CLID").ToString
                    LabelData.CHCode = ds.Tables(0).Rows(0)("ItemText").ToString
                    LabelData.Qty = Format(ds.Tables(0).Rows(0)("QtyBaseUOM"), "#0.#####")
                    LabelData.DateCode = FixNull(ds.Tables(0).Rows(0)("DateCode"))
                    LabelData.LotNo = FixNull(ds.Tables(0).Rows(0)("LotNo"))
                    LabelData.VendorID = FixNull(ds.Tables(0).Rows(0)("VendorID"))

                    Dim ItemList(1) As String
                    ItemList(0) = ds.Tables(0).Rows(0)("MaterialNo").ToString
                    Dim dsAML As New DataSet
                    Try
                        dsAML = GetAML(ItemList)
                    Catch ex As Exception
                    End Try

                    If dsAML Is Nothing OrElse dsAML.Tables.Count = 0 OrElse dsAML.Tables("ItemData").Rows.Count = 0 Then
                    ElseIf dsAML.Tables("ItemData").Rows.Count > 0 Then
                        If Mid(ds.Tables(0).Rows(0)("MaterialNo").ToString, 1, 3) = "509" Then
                            LabelData.Rev = Split(dsAML.Tables("ItemData").Rows(0)("RESData"), "~")(0)
                        Else
                            LabelData.Rev = ""
                        End If
                        LabelData.MSL = FixNull(dsAML.Tables("ItemData").Rows(0)("MSL"))
                    End If

                    Dim strContent As String
                    Dim arryFile() As String

                    arryFile = Split(CH09LabelFile, "\")
                    strContent = "CHCode^" & LabelData.CHCode & "^Qty^" & LabelData.Qty & "^VendorID^" & LabelData.VendorID & "^DateCode^" & LabelData.DateCode & "^LotNo^" & LabelData.LotNo & "^Rev^" & LabelData.Rev & "^MSL^" & LabelData.MSL

                    Sqlstr = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                    da.ExecuteScalar(Sqlstr)
                    Sleep(5)
                Next

            Catch ex As Exception
                ErrorLogging("PublicFunction-PrintCH09Label", "", ex.Message & ex.Source, "E")
                PrintCH09Label = False
            End Try
        End Using

    End Function


#End Region

#Region "OraclePublic"
    Public Function AutoMail_SiplaceDataCheck() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim flag As Boolean = True
            sda = da.Sda_Sele()
            Try
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_RptSiplaceDataCheck"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Dim folderpath As String
                    folderpath = "D:\AutoReportV2\ReportsFiles\SiplaceDataCheck"
                    Dim fileName As String
                    Dim fileNamepath As String
                    Dim sheetName As String
                    fileName = "SiplaceDataCheck" & Format(DateTime.Now, "yyyyMMddHHmmss") & ".xlsx"
                    fileNamepath = folderpath & "\" & fileName
                    sheetName = "SiplaceDataCheck"
                    GenerateExcel(ds.Tables(0), fileNamepath, sheetName)
                    Dim mailsubject As String
                    Dim mailtext As String
                    Dim mailattch As String
                    'Dim mailto() As String = {"charleszhao@artesyn.com"}
                    Dim mailto() As String = {"SunnyLau@artesyn.com", "RachelYu@artesyn.com", "wuchengqin@artesyn.com", "charleszhao@artesyn.com", "Shihong.Lin@artesyn.com", "Kent.Yang@artesyn.com"}
                    mailsubject = "Alert-Some SMT boards missing Siplace data"
                    mailtext = "<font size='3' color='red'>Hi,ALL</font><br><font size='3' color='black'>"
                    mailtext = mailtext & "Please check below file for SMT boards missing Siplace data . Thanks</font>"
                    mailtext = mailtext & "<br><br><font size='3' color='blue'>\\cnapgzhofs10\ReportsFiles\SiplaceDataCheck\" & fileName & "</font>"
                    mailattch = ""
                    SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext, mailattch)
                    flag = True
                Else
                    flag = False
                End If
                Return flag
            Catch oe As Exception
                ErrorLogging("SiplaceDataCheck", "Background", oe.Message & oe.Source, "E")
                Return False
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function AutoSendMail_PastDueDJ() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim mods As New DataSet
            Dim l_count As Integer
            sda = da.Sda_Sele()
            Try
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_bkg_sendmail_pastduedj"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                Dim folderpath As String
                folderpath = "D:\AutoEJITErrorReport\ZS_PastDueDJ_" & Format(DateTime.Now, "yyyyMMdd")
                Directory.CreateDirectory(folderpath)

                Dim fileName As String
                Dim fileNamepath As String
                Dim sheetName As String
                fileName = "ZS_PastdueDj_" & Format(DateTime.Now, "yyyyMMdd") & "_Release.xlsx"
                fileNamepath = folderpath & "\" & fileName
                sheetName = "Data"
                GenerateExcel(ds.Tables(0), fileNamepath, sheetName)
                fileName = "ZS_PastdueDj_" & Format(DateTime.Now, "yyyyMMdd") & "_Unrelease.xlsx"
                fileNamepath = folderpath & "\" & fileName
                sheetName = "Data"
                GenerateExcel(ds.Tables(1), fileNamepath, sheetName)

                Dim zipPath As String = folderpath & ".zip"
                ZipFile.CreateFromDirectory(folderpath, zipPath, CompressionLevel.Fastest, True)

                Dim mailsubject As String
                Dim mailtext As String
                Dim mailattch As String
                Dim mailto() As String = {"SuriZhang@artesyn.com", "SummerChan@artesyn.com", "ClaraMo@artesyn.com", "EricChan@artesyn.com", "CatherineYip@artesyn.com", "SunnyLau@artesyn.com", "janecao@artesyn.com", "MaggieMo@artesyn.com", "charleszhao@artesyn.com"}
                'Dim mailto() As String = {"charleszhao@artesyn.com"}

                mailsubject = "Past Due DJ as of " & Format(DateTime.Now, "MMM").ToString & " " & DateTime.Now.Day.ToString & " " & DateTime.Now.DayOfWeek.ToString
                mailtext = ""
                'mailtext = "<font size='4' color='red'>Hi,ALL</font><br><font size='3' color='black'>"
                'mailtext = mailtext & "Send Mail Test - Past Due DJ . Thanks</font>"
                'mailtext = mailtext & "<br><br><font size='4' color='blue'>\\CNAPGZHOETDV01\Temp\AutoEJITLogFile\" & fileName & "</font>"
                mailattch = zipPath

                SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext, mailattch)

                Return True
            Catch oe As Exception
                ErrorLogging("AutoSendMail", "AutoSendMail_PastDueDJ", oe.Message & oe.Source, "E")
                Return False
            Finally
                If l_count > 0 Then
                    If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                End If
            End Try
        End Using
    End Function
    Public Function DeleteMOAllocated(ByVal delMOItemList As DataSet, ByVal ERPLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim j As Integer
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim dr() As DataRow
            Dim strsql As String
            Dim mo As String
            Try
                sda = da.Sda_Sele()
                oda = da.Oda_Insert()
                If Not delMOItemList Is Nothing AndAlso delMOItemList.Tables.Count > 0 AndAlso delMOItemList.Tables(0).Rows.Count > 0 Then
                    mo = delMOItemList.Tables(0).Rows(0)("MO")
                    strsql = String.Format("Delete T_DeleteMOAllocated where ORGID = {0} and MO = '{1}'", ERPLoginData.OrgID, mo)
                    da.ExecuteNonQuery(strsql)
                    For i = 0 To delMOItemList.Tables(0).Rows.Count - 1
                        strsql = String.Format("insert into T_DeleteMOAllocated(ORGID , MO , Item) values ({0},'{1}','{2}')", ERPLoginData.OrgID, delMOItemList.Tables(0).Rows(i)("MO"), delMOItemList.Tables(0).Rows(i)("ITEM"))
                        da.ExecuteNonQuery(strsql)
                    Next
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_del_ledmoallocated"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@mo_num", SqlDbType.VarChar, 250).Value = mo
                    sda.SelectCommand.Parameters.Add("@org_id", SqlDbType.Int).Value = ERPLoginData.OrgID
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(ds)
                    sda.SelectCommand.Connection.Close()
                    For j = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(j).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(j).SetAdded()
                        End If
                    Next
                    oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_wip_pkg.delete_allocations"
                    oda.InsertCommand.Parameters.Add("p_mo_line_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Parameters("p_mo_line_id").SourceColumn = "line_id"
                    oda.InsertCommand.Parameters("o_success_flag").SourceColumn = "o_flag"
                    oda.InsertCommand.Parameters("o_error_mssg").SourceColumn = "o_mssg"
                    oda.InsertCommand.Connection.Open()
                    oda.Update(ds.Tables(0))
                    oda.InsertCommand.Connection.Close()
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        For j = 0 To delMOItemList.Tables(0).Rows.Count - 1
                            If ds.Tables(0).Rows(i)("mo") = delMOItemList.Tables(0).Rows(j)("mo") And ds.Tables(0).Rows(i)("item") = delMOItemList.Tables(0).Rows(j)("item") Then
                                delMOItemList.Tables(0).Rows(j)("O_FLAG") = ds.Tables(0).Rows(i)("o_flag")
                                delMOItemList.Tables(0).Rows(j)("O_MSSG") = ds.Tables(0).Rows(i)("o_mssg")
                            End If
                        Next
                        delMOItemList.Tables(0).AcceptChanges()
                    Next
                End If
                Return delMOItemList
            Catch oe As Exception
                ErrorLogging("DeleteMOAllocated", "eTrace_WS_Background", oe.Message & oe.Source, "E")
                Return delMOItemList
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function RedoLedMOAllocated(ByVal EventID As String, ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim MOStatus As String
            RedoLedMOAllocated = ""
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_redo_ledmoallocated"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@eventid", SqlDbType.VarChar, 250).Value = EventID
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(2).Rows.Count > 0 Then
                    MOStatus = ds.Tables(2).Rows(0)("OraMOStatus")
                End If
                oda = da.Oda_Insert()
                If MOStatus = "Closed" Then
                    RedoLedMOAllocated = "Y"                                             'Oracle is already successful
                    Return RedoLedMOAllocated
                Else
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(i).SetAdded()
                        End If
                    Next
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_wip_pkg.delete_allocations"
                    oda.InsertCommand.Parameters.Add("p_mo_line_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Parameters("p_mo_line_id").SourceColumn = "line_id"
                    oda.InsertCommand.Parameters("o_success_flag").SourceColumn = "o_flag"
                    oda.InsertCommand.Parameters("o_error_mssg").SourceColumn = "o_mssg"
                    oda.InsertCommand.Connection.Open()
                    oda.Update(ds.Tables(0))
                    oda.InsertCommand.Connection.Close()
                    Dim dr() As DataRow = Nothing
                    dr = ds.Tables(0).Select("o_flag = 'N'")
                    If dr.Length > 0 Then
                        RedoLedMOAllocated = dr(0)("o_mssg").ToString
                        Return RedoLedMOAllocated
                    End If
                    For i = 0 To ds.Tables(1).Rows.Count - 1
                        If ds.Tables(1).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(1).Rows(i).SetAdded()
                        End If
                    Next
                    oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_wip_pkg.create_ledmo_allocation"
                    oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_moheader_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_item_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_subinv", OracleType.VarChar, 250)
                    oda.InsertCommand.Parameters.Add("p_locator_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_lot_num", OracleType.VarChar, 150)
                    oda.InsertCommand.Parameters.Add("p_qty", OracleType.Double)
                    oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)
                    oda.InsertCommand.Parameters("p_org_id").Value = ERPLoginData.OrgID
                    oda.InsertCommand.Parameters("p_moheader_id").SourceColumn = "header_id"
                    oda.InsertCommand.Parameters("p_item_id").SourceColumn = "item_id"
                    oda.InsertCommand.Parameters("p_subinv").SourceColumn = "subinv"
                    oda.InsertCommand.Parameters("p_locator_id").SourceColumn = "locator_id"
                    oda.InsertCommand.Parameters("p_lot_num").SourceColumn = "lot_num"
                    oda.InsertCommand.Parameters("p_qty").SourceColumn = "qty"
                    oda.InsertCommand.Parameters("p_user_id").Value = ERPLoginData.UserID
                    oda.InsertCommand.Parameters("p_resp_id").Value = ERPLoginData.RespID_Inv
                    oda.InsertCommand.Parameters("p_appl_id").Value = ERPLoginData.AppID_Inv
                    oda.InsertCommand.Parameters("o_success_flag").SourceColumn = "o_flag"
                    oda.InsertCommand.Parameters("o_error_mssg").SourceColumn = "o_mssg"
                    oda.InsertCommand.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput
                    oda.InsertCommand.Connection.Open()
                    oda.Update(ds.Tables(1))
                    oda.InsertCommand.Connection.Close()

                    Dim dr1() As DataRow = Nothing
                    dr1 = ds.Tables(1).Select("o_flag = 'N'")
                    If dr1.Length > 0 Then
                        RedoLedMOAllocated = "Item " & dr1(0)("MaterialNo") & ", " & dr1(0)("o_mssg").ToString
                        ErrorLogging("RedoLedMOAllocated1", ERPLoginData.User, "MO " & dr1(0)("MO") & ", " & RedoLedMOAllocated, "I")
                        RedoLedMOAllocated = RedoLedMOAllocated.Replace("create_ledmo_allocation_err_code", "ErrCode")
                        Return RedoLedMOAllocated
                    Else
                        RedoLedMOAllocated = "R"                                'Need to submit data to Oracle again
                    End If
                End If
            Catch oe As Exception
                RedoLedMOAllocated = "N"
                ErrorLogging("RedoLedMOAllocated", ERPLoginData.User, oe.Message & oe.Source, "E")
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Private Sub GenerateExcel(ByVal dataToExcel As System.Data.DataTable, ByVal fileName As String, ByVal sheetName As String)
        Dim currentDirectorypath As String = Environment.CurrentDirectory
        Dim finalFileNameWithPath As String = String.Empty
        finalFileNameWithPath = fileName ''String.Format("{0}\{1}.xlsx", currentDirectorypath, fileName)
        Dim newFile As System.IO.FileInfo = New FileInfo(finalFileNameWithPath)
        Using pck As New OfficeOpenXml.ExcelPackage(newFile)
            Dim ws As OfficeOpenXml.ExcelWorksheet = pck.Workbook.Worksheets.Add(sheetName)
            ws.Cells("A1").LoadFromDataTable(dataToExcel, True)
            ws.Cells.Style.Font.Size = 9
            ws.Cells.AutoFitColumns()
            ws.Cells("1:1").Style.Font.Bold = True
            ws.Cells("1:1").Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            ws.Cells("H:H").Style.Numberformat.Format = "YYYY-MM-DD"
            pck.Save()
        End Using
    End Sub
    Private Function FileExists(ByVal FileFullPath As String) As Boolean
        If Trim(FileFullPath) = "" Then
            Return False
        End If
        Dim f As New IO.FileInfo(FileFullPath)
        Return f.Exists
    End Function
    Public Function SendMail(ByVal mailfrom As String, ByVal mailto() As String, ByVal mailsubject As String, ByVal message As String, Optional ByVal attachmentfile As String = "") As Boolean
        Try
            Dim smtpclient As New SmtpClient()
            Dim mail As New MailMessage()
            smtpclient.EnableSsl = False
            smtpclient.UseDefaultCredentials = True
            smtpclient.Port = 25
            smtpclient.Host = "ECP-GATEWAY.ASTEC-POWER.COM"
            mail = New MailMessage()
            mail.IsBodyHtml = True
            mail.From = New MailAddress(mailfrom)
            For Each value As String In mailto
                mail.To.Add(value)
            Next
            mail.Subject = mailsubject
            mail.Body = message
            If FileExists(attachmentfile) Then
                Dim attachment As Attachment
                attachment = New Attachment(attachmentfile)
                mail.Attachments.Add(attachment)
            End If
            smtpclient.Send(mail)
            Return True
        Catch oe As Exception
            Return False
        End Try
    End Function
    Public Function AutoCreatedEJIT() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim mods As New DataSet
            Dim moorgcode As New DataSet
            Dim i As Integer
            Dim j As Integer
            Dim scrow As DataRow
            Dim user_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_User_id_ODay")
            Dim resp_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_Resp_id")
            Dim wfoe_resp_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_WFOE_Resp_id")
            Dim appl_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_Appl_id")
            Dim Sqlstr As String
            Dim l_count As Integer
            Dim Batch_id As Integer
            Dim error_flag As String
            Dim CheckStatus As String
            Dim uploadflag As Boolean
            sda = da.Sda_Sele()
            oda = da.Oda_Insert_EJIT()
            Try
                'Sqlstr = "select "
                'Sqlstr = Sqlstr & " (case when "
                'Sqlstr = Sqlstr & " (select count(1) from T_AutoeJIT_RunningStatus where convert(date,convert(char(10),StartTime,120),120) = convert(date,convert(char(10),eTrace.dbo.ora_runippdate(),120),120)) = 0 then 1"
                'Sqlstr = Sqlstr & " else"
                'Sqlstr = Sqlstr & " (select count(1) from T_AutoeJIT_RunningStatus where CalculateFlag = 'N' and "
                'Sqlstr = Sqlstr & " StartTime = isnull((select max(StartTime) from T_AutoeJIT_RunningStatus where convert(date,convert(char(10),StartTime,120),120) = convert(date,convert(char(10),eTrace.dbo.ora_runippdate(),120),120)),(select max(StartTime) from T_AutoeJIT_RunningStatus)))"
                'Sqlstr = Sqlstr & " end) l_count"
                'l_count = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_AutoEJITStatusCheck"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@Validflag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                CheckStatus = sda.SelectCommand.Parameters("@Validflag").Value.ToString()
                If CheckStatus = "N" Then
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 AutoEJIT Status check error", "E")
                    Return False
                    Exit Function
                End If
                If CheckStatus = "Y" Then
                    'user_id = 21745   '21745 KB_SCHEDULER 21746 KB_SCHEDULER_OFF
                    ''resp_id = 55768  'ZS
                    'resp_id = 57109   'LD
                    uploadflag = UploadAutoEJIT_DJSum_to_Oracle()
                    If uploadflag = False Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 Upload DJ Sum to oracle error", "I")
                        Return False
                        Exit Function
                    End If
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_autoejit_calculate_ipp"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.SelectCommand.ExecuteNonQuery()
                    sda.SelectCommand.Connection.Close()

                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_autoejit_calculate_ipp error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    uploadflag = UploadAutoEJIT_DJList_to_Oracle()
                    If uploadflag = False Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 Upload DJ List to oracle error", "I")
                        Return False
                        Exit Function
                    End If
                    ds.Tables.Clear()
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_auto_ejit"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(ds)
                    sda.SelectCommand.Connection.Close()
                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_auto_ejit error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(i).SetAdded()
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step1 calculate upload data", "I")
                    'oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.insert_pr_interface"
                    oda.InsertCommand.Parameters.Add("P_BATCH_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_GROUP_CODE", OracleType.VarChar, 30)
                    oda.InsertCommand.Parameters.Add("P_LINE_TYPE_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_AUTHORIZATION_STATUS", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_PREPARER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_ITEM_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_QUANTITY", OracleType.Double)
                    oda.InsertCommand.Parameters.Add("P_NEED_BY_DATE", OracleType.DateTime)
                    oda.InsertCommand.Parameters.Add("P_DELIVER_TO_LOCATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DELIVER_TO_REQUESTOR_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_ORG_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_REQUISITION_TYPE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_ORGANIZATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_SUBINVENTORY", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_SUBINVENTORY", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("P_NOTE_TO_RECEIVER", OracleType.VarChar, 240)
                    oda.InsertCommand.Parameters.Add("P_HEADER_DESCRIPTION", OracleType.VarChar, 240)
                    oda.InsertCommand.Parameters.Add("P_CHARGE_ACCOUNT_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_ORGANIZATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_INTERFACE_SOURCE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_FLAG", OracleType.VarChar, 1)
                    oda.InsertCommand.Parameters.Add("P_DOCUMENT_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_DOC_HEADER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_DOC_LINE_NUM", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_VENDOR_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_VENDOR_SITE_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_BUYER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_CREATED_BY", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_LAST_UPDATED_BY", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_JUSTIFICATION", OracleType.VarChar, 500)
                    oda.InsertCommand.Parameters("P_BATCH_ID").SourceColumn = "batch_id"
                    oda.InsertCommand.Parameters("P_GROUP_CODE").SourceColumn = "group_code"
                    oda.InsertCommand.Parameters("P_LINE_TYPE_ID").SourceColumn = "line_type_id"
                    oda.InsertCommand.Parameters("P_AUTHORIZATION_STATUS").SourceColumn = "authorization_status"
                    oda.InsertCommand.Parameters("P_PREPARER_ID").SourceColumn = "preparer_id"
                    oda.InsertCommand.Parameters("P_ITEM_ID").SourceColumn = "item_id"
                    oda.InsertCommand.Parameters("P_QUANTITY").SourceColumn = "quantity"
                    oda.InsertCommand.Parameters("P_NEED_BY_DATE").SourceColumn = "need_by_date"
                    oda.InsertCommand.Parameters("P_DELIVER_TO_LOCATION_ID").SourceColumn = "deliver_to_location_id"
                    oda.InsertCommand.Parameters("P_DELIVER_TO_REQUESTOR_ID").SourceColumn = "deliver_to_requestor_id"
                    oda.InsertCommand.Parameters("P_ORG_ID").SourceColumn = "org_id"
                    oda.InsertCommand.Parameters("P_REQUISITION_TYPE").SourceColumn = "requisition_type"
                    oda.InsertCommand.Parameters("P_SOURCE_ORGANIZATION_ID").SourceColumn = "source_organization_id"
                    oda.InsertCommand.Parameters("P_SOURCE_SUBINVENTORY").SourceColumn = "source_subinventory"
                    oda.InsertCommand.Parameters("P_SOURCE_TYPE_CODE").SourceColumn = "source_type_code"
                    oda.InsertCommand.Parameters("P_DESTINATION_SUBINVENTORY").SourceColumn = "destination_subinventory"
                    oda.InsertCommand.Parameters("P_NOTE_TO_RECEIVER").SourceColumn = "note_to_receiver"
                    oda.InsertCommand.Parameters("P_HEADER_DESCRIPTION").SourceColumn = "header_description"
                    oda.InsertCommand.Parameters("P_CHARGE_ACCOUNT_ID").SourceColumn = "charge_account_id"
                    oda.InsertCommand.Parameters("P_DESTINATION_ORGANIZATION_ID").SourceColumn = "destination_organization_id"
                    oda.InsertCommand.Parameters("P_DESTINATION_TYPE_CODE").SourceColumn = "destination_type_code"
                    oda.InsertCommand.Parameters("P_INTERFACE_SOURCE_CODE").SourceColumn = "interface_source_code"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_FLAG").SourceColumn = "autosource_flag"
                    oda.InsertCommand.Parameters("P_DOCUMENT_TYPE_CODE").SourceColumn = "document_type_code"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_DOC_HEADER_ID").SourceColumn = "autosource_doc_header_id"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_DOC_LINE_NUM").SourceColumn = "autosource_doc_line_num"
                    oda.InsertCommand.Parameters("P_SUGGESTED_VENDOR_ID").SourceColumn = "suggested_vendor_id"
                    oda.InsertCommand.Parameters("P_SUGGESTED_VENDOR_SITE_ID").SourceColumn = "suggested_vendor_site_id"
                    oda.InsertCommand.Parameters("P_SUGGESTED_BUYER_ID").SourceColumn = "suggested_buyer_id"
                    oda.InsertCommand.Parameters("P_CREATED_BY").SourceColumn = "created_by"
                    oda.InsertCommand.Parameters("P_LAST_UPDATED_BY").SourceColumn = "last_updated_by"
                    oda.InsertCommand.Parameters("P_JUSTIFICATION").SourceColumn = "justification"
                    oda.InsertCommand.Connection.Open()
                    oda.Update(ds.Tables(0))
                    oda.InsertCommand.Connection.Close()
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step2 insert data to interface", "I")
                    oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.del_po_requisitions_inte"
                    oda.InsertCommand.Parameters.Add("o_return_message", OracleType.VarChar, 1000)
                    oda.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("p_created_by", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_date_from", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_date_to", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_interface_source_code", OracleType.VarChar, 250)
                    oda.InsertCommand.Parameters("o_return_message").Direction = ParameterDirection.Output
                    oda.InsertCommand.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                    oda.InsertCommand.Parameters("p_created_by").Value = ConfigurationManager.AppSettings("AutoEJIT_User_id_ODay")
                    oda.InsertCommand.Parameters("p_date_from").Value = Format(DateTime.Now, "yyyy/MM/dd") & " 00:00:01"
                    oda.InsertCommand.Parameters("p_date_to").Value = Format(DateTime.Now, "yyyy/MM/dd") & " 23:59:59"
                    oda.InsertCommand.Parameters("p_interface_source_code").Value = "SKIPNEXTDATE"
                    oda.InsertCommand.Connection.Open()
                    oda.InsertCommand.ExecuteNonQuery()
                    oda.InsertCommand.Connection.Close()
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_autoejit_save"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.SelectCommand.ExecuteNonQuery()
                    sda.SelectCommand.Connection.Close()
                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_autoejit_save error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step3 save upload data to EJIT table", "I")
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_get_autoejit_mo_orgcode"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(moorgcode)
                    sda.SelectCommand.Connection.Close()
                    For j = 0 To moorgcode.Tables(0).Rows.Count - 1
                        sda.SelectCommand.Parameters.Clear()
                        mods.Clear()
                        sda.SelectCommand.CommandType = CommandType.StoredProcedure
                        sda.SelectCommand.CommandText = "ora_get_autoejit_mo"
                        sda.SelectCommand.CommandTimeout = TimeOut_M5
                        sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 250).Value = moorgcode.Tables(0).Rows(j)("ORG_CODE")
                        sda.SelectCommand.Connection.Open()
                        sda.Fill(mods)
                        sda.SelectCommand.Connection.Close()
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step4 get MO data", "I")
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.get_mo_batch_id"
                        oda.InsertCommand.Parameters.Add("o_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_batch_id").Direction = ParameterDirection.Output
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()

                        Batch_id = oda.InsertCommand.Parameters("o_batch_id").Value
                        For i = 0 To mods.Tables(0).Rows.Count - 1
                            If mods.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                                mods.Tables(0).Rows(i).SetAdded()
                            End If
                        Next
                        For i = 0 To mods.Tables(0).Rows.Count - 1
                            mods.Tables(0).Rows(i)("BATCH_ID") = Batch_id
                            mods.Tables(0).Rows(i)("USERID") = user_id
                        Next
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.insert_mo_data"
                        oda.InsertCommand.Parameters.Add("p_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                        oda.InsertCommand.Parameters.Add("p_mo_qty", OracleType.Double)
                        oda.InsertCommand.Parameters.Add("p_subinv", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_reference", OracleType.VarChar, 240)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                        oda.InsertCommand.Parameters("p_batch_id").SourceColumn = "batch_id"
                        oda.InsertCommand.Parameters("p_item_num").SourceColumn = "item_no"
                        oda.InsertCommand.Parameters("p_mo_qty").SourceColumn = "mo_qty"
                        oda.InsertCommand.Parameters("p_subinv").SourceColumn = "subinv"
                        oda.InsertCommand.Parameters("p_locator").SourceColumn = "locator"
                        oda.InsertCommand.Parameters("p_reference").SourceColumn = "reference"
                        oda.InsertCommand.Parameters("p_user_id").SourceColumn = "userid"
                        oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                        oda.InsertCommand.Parameters("o_msg").SourceColumn = "o_msg"
                        oda.InsertCommand.Connection.Open()
                        oda.Update(mods.Tables(0))
                        oda.InsertCommand.Connection.Close()
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.autoejit_mo_submit_req"
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 1000)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_msg").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("p_org_code").Value = moorgcode.Tables(0).Rows(j)("ORG_CODE")
                        oda.InsertCommand.Parameters("p_batch_id").Value = Batch_id
                        oda.InsertCommand.Parameters("p_user_id").Value = user_id
                        If moorgcode.Tables(0).Rows(j)("ORG_CODE") = "680" Then
                            oda.InsertCommand.Parameters("p_resp_id").Value = wfoe_resp_id
                        Else
                            oda.InsertCommand.Parameters("p_resp_id").Value = resp_id
                        End If
                        oda.InsertCommand.Parameters("p_appl_id").Value = appl_id
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()

                        error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                        If error_flag <> "Y" Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "E")
                        End If
                        If Len(oda.InsertCommand.Parameters("o_msg").Value.ToString()) > 0 Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "I")
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step5 submit request to create MO", "I")
                    'Dim SourceCodeData As System.Data.DataTable
                    'Dim myDataColumn As DataColumn
                    'SourceCodeData = New Data.DataTable("InterfaceSourceCode")
                    'myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
                    'SourceCodeData.Columns.Add(myDataColumn)
                    'myDataColumn = New Data.DataColumn("SourceCode", System.Type.GetType("System.String"))
                    'SourceCodeData.Columns.Add(myDataColumn)
                    'ds.Tables.Add(SourceCodeData)
                    'scrow = ds.Tables("InterfaceSourceCode").NewRow
                    'scrow("SourceCode") = "AutoEJIT1530"
                    'scrow("OrgCode") = "580"
                    'scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_EJIT").ToString
                    'ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    'ds.Tables("InterfaceSourceCode").AcceptChanges()
                    'scrow = ds.Tables("InterfaceSourceCode").NewRow
                    'scrow("SourceCode") = "AUTOHUB1530"
                    'scrow("OrgCode") = "580"
                    'scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_HUB").ToString
                    'ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    'ds.Tables("InterfaceSourceCode").AcceptChanges()
                    'If ConfigurationManager.AppSettings("AutoEJIT_SC_IR").ToString <> "NO" Then
                    '    scrow = ds.Tables("InterfaceSourceCode").NewRow
                    '    scrow("OrgCode") = "580"
                    '    scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_IR").ToString
                    '    ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    '    ds.Tables("InterfaceSourceCode").AcceptChanges()
                    ds.Tables(1).TableName = "InterfaceSourceCode"
                    For i = 0 To ds.Tables("InterfaceSourceCode").Rows.Count - 1
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.autoejit_po_submit_req"
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 1000)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_source_code", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_msg").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("p_source_code").Value = ds.Tables("InterfaceSourceCode").Rows(i)("SourceCode")
                        oda.InsertCommand.Parameters("p_user_id").Value = user_id
                        If ds.Tables("InterfaceSourceCode").Rows(i)("OrgCode") = "680" Then
                            oda.InsertCommand.Parameters("p_resp_id").Value = wfoe_resp_id
                        Else
                            oda.InsertCommand.Parameters("p_resp_id").Value = resp_id
                        End If
                        oda.InsertCommand.Parameters("p_appl_id").Value = appl_id
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()
                        error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                        If error_flag <> "Y" Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "E")
                        End If
                        If Len(oda.InsertCommand.Parameters("o_msg").Value.ToString()) > 0 Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "I")
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step6 submit request to create PO/IR", "I")
                    Dim mailtype As Integer
                    Sqlstr = "select dbo.ora_autoejit_mailtype() mail_type"
                    mailtype = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                    ds.Tables.Clear()
                    Dim dt As New DataTable
                    Dim fileName As String
                    Dim fileNamepath As String
                    Dim sheetName As String
                    fileName = Format(DateTime.Now, "yyyyMMddHHmmss") & ".xlsx"
                    fileNamepath = "D:\AutoEJITErrorReport\" & fileName
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_get_autoejit_errorlist"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(ds)
                    sda.SelectCommand.Connection.Close()
                    sheetName = "Sheet1"
                    GenerateExcel(ds.Tables(0), fileNamepath, sheetName)
                    Dim mailsubject As String
                    Dim mailtext As String
                    Dim mailto() As String = {"charleszhao@artesyn.com", "janecao@artesyn.com", "MaggieMo@artesyn.com"}
                    Dim mailtouser() As String = {"charleszhao@artesyn.com", "janecao@artesyn.com", "MaggieMo@artesyn.com", ConfigurationManager.AppSettings("AutoEJIT_ErrMailGroup").ToString}
                    If mailtype = -1 Then
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram calculate error"
                        mailtext = "Failed"
                        SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext)
                    ElseIf mailtype = 0 Then
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram Error Message records list"
                        mailtext = "<font size='4' color='red'>Hi,ALL</font><br><font size='3' color='black'>"
                        mailtext = mailtext & "Please click the below path to check the error message for AutoEJIT program . Thanks</font>"
                        mailtext = mailtext & "<br><br><font size='4' color='blue'>\\CNAPGZHOFS10\AutoEJITErrorReport\" & fileName & "</font>"
                        SendMail("cnapgzhofs10@artesyn.com", mailtouser, mailsubject, mailtext)
                    Else
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram run successfully"
                        mailtext = "Successful"
                        SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext)
                    End If
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step7 Send mail via webservice", "I")
                End If
                Return True
            Catch oe As Exception
                ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oe.Message & oe.Source, "E")
                Return False
            Finally
                If l_count > 0 Then
                    If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                    If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
                End If
            End Try
        End Using
    End Function
    Public Function AutoCreatedEJIT_ByDay() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim mods As New DataSet
            Dim moorgcode As New DataSet
            Dim i As Integer
            Dim j As Integer
            Dim scrow As DataRow
            Dim user_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_User_id_WDay")
            Dim resp_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_Resp_id")
            Dim wfoe_resp_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_WFOE_Resp_id")
            Dim appl_id As Integer = ConfigurationManager.AppSettings("AutoEJIT_Appl_id")
            Dim Sqlstr As String
            Dim l_count As Integer
            Dim Batch_id As Integer
            Dim error_flag As String
            Dim CheckStatus As String
            Dim uploadflag As Boolean
            sda = da.Sda_Sele()
            oda = da.Oda_Insert_EJIT()
            Try
                'Sqlstr = "select "
                'Sqlstr = Sqlstr & " (case when "
                'Sqlstr = Sqlstr & " (select count(1) from T_AutoeJIT_RunningStatus where convert(date,convert(char(10),StartTime,120),120) = convert(date,convert(char(10),eTrace.dbo.ora_runippdate(),120),120)) = 0 then 1"
                'Sqlstr = Sqlstr & " else"
                'Sqlstr = Sqlstr & " (select count(1) from T_AutoeJIT_RunningStatus where CalculateFlag = 'N' and "
                'Sqlstr = Sqlstr & " StartTime = isnull((select max(StartTime) from T_AutoeJIT_RunningStatus where convert(date,convert(char(10),StartTime,120),120) = convert(date,convert(char(10),eTrace.dbo.ora_runippdate(),120),120)),(select max(StartTime) from T_AutoeJIT_RunningStatus)))"
                'Sqlstr = Sqlstr & " end) l_count"
                'l_count = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_AutoEJITStatusCheck"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@Validflag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                CheckStatus = sda.SelectCommand.Parameters("@Validflag").Value.ToString()
                If CheckStatus = "N" Then
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 AutoEJIT Status check error", "E")
                    Return False
                    Exit Function
                End If
                If CheckStatus = "Y" Then
                    'user_id = 21745   '21745 KB_SCHEDULER 21746 KB_SCHEDULER_OFF
                    ''resp_id = 55768  'ZS
                    'resp_id = 57109   'LD
                    uploadflag = UploadAutoEJIT_DJSum_to_Oracle()
                    If uploadflag = False Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 Upload DJ Sum to oracle error", "I")
                        Return False
                        Exit Function
                    End If
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_autoejit_calculate_ipp"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.SelectCommand.ExecuteNonQuery()
                    sda.SelectCommand.Connection.Close()

                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_autoejit_calculate_ipp error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    uploadflag = UploadAutoEJIT_DJList_to_Oracle()
                    If uploadflag = False Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step0 Upload DJ List to oracle error", "I")
                        Return False
                        Exit Function
                    End If
                    ds.Tables.Clear()
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_auto_ejit"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(ds)
                    sda.SelectCommand.Connection.Close()
                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_auto_ejit error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(i).SetAdded()
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step1 calculate upload data", "I")
                    'oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.insert_pr_interface"
                    oda.InsertCommand.Parameters.Add("P_BATCH_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_GROUP_CODE", OracleType.VarChar, 30)
                    oda.InsertCommand.Parameters.Add("P_LINE_TYPE_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_AUTHORIZATION_STATUS", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_PREPARER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_ITEM_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_QUANTITY", OracleType.Double)
                    oda.InsertCommand.Parameters.Add("P_NEED_BY_DATE", OracleType.DateTime)
                    oda.InsertCommand.Parameters.Add("P_DELIVER_TO_LOCATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DELIVER_TO_REQUESTOR_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_ORG_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_REQUISITION_TYPE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_ORGANIZATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_SUBINVENTORY", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("P_SOURCE_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_SUBINVENTORY", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("P_NOTE_TO_RECEIVER", OracleType.VarChar, 240)
                    oda.InsertCommand.Parameters.Add("P_HEADER_DESCRIPTION", OracleType.VarChar, 240)
                    oda.InsertCommand.Parameters.Add("P_CHARGE_ACCOUNT_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_ORGANIZATION_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_DESTINATION_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_INTERFACE_SOURCE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_FLAG", OracleType.VarChar, 1)
                    oda.InsertCommand.Parameters.Add("P_DOCUMENT_TYPE_CODE", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_DOC_HEADER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_AUTOSOURCE_DOC_LINE_NUM", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_VENDOR_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_VENDOR_SITE_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_SUGGESTED_BUYER_ID", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_CREATED_BY", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_LAST_UPDATED_BY", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("P_JUSTIFICATION", OracleType.VarChar, 500)
                    oda.InsertCommand.Parameters("P_BATCH_ID").SourceColumn = "batch_id"
                    oda.InsertCommand.Parameters("P_GROUP_CODE").SourceColumn = "group_code"
                    oda.InsertCommand.Parameters("P_LINE_TYPE_ID").SourceColumn = "line_type_id"
                    oda.InsertCommand.Parameters("P_AUTHORIZATION_STATUS").SourceColumn = "authorization_status"
                    oda.InsertCommand.Parameters("P_PREPARER_ID").SourceColumn = "preparer_id"
                    oda.InsertCommand.Parameters("P_ITEM_ID").SourceColumn = "item_id"
                    oda.InsertCommand.Parameters("P_QUANTITY").SourceColumn = "quantity"
                    oda.InsertCommand.Parameters("P_NEED_BY_DATE").SourceColumn = "need_by_date"
                    oda.InsertCommand.Parameters("P_DELIVER_TO_LOCATION_ID").SourceColumn = "deliver_to_location_id"
                    oda.InsertCommand.Parameters("P_DELIVER_TO_REQUESTOR_ID").SourceColumn = "deliver_to_requestor_id"
                    oda.InsertCommand.Parameters("P_ORG_ID").SourceColumn = "org_id"
                    oda.InsertCommand.Parameters("P_REQUISITION_TYPE").SourceColumn = "requisition_type"
                    oda.InsertCommand.Parameters("P_SOURCE_ORGANIZATION_ID").SourceColumn = "source_organization_id"
                    oda.InsertCommand.Parameters("P_SOURCE_SUBINVENTORY").SourceColumn = "source_subinventory"
                    oda.InsertCommand.Parameters("P_SOURCE_TYPE_CODE").SourceColumn = "source_type_code"
                    oda.InsertCommand.Parameters("P_DESTINATION_SUBINVENTORY").SourceColumn = "destination_subinventory"
                    oda.InsertCommand.Parameters("P_NOTE_TO_RECEIVER").SourceColumn = "note_to_receiver"
                    oda.InsertCommand.Parameters("P_HEADER_DESCRIPTION").SourceColumn = "header_description"
                    oda.InsertCommand.Parameters("P_CHARGE_ACCOUNT_ID").SourceColumn = "charge_account_id"
                    oda.InsertCommand.Parameters("P_DESTINATION_ORGANIZATION_ID").SourceColumn = "destination_organization_id"
                    oda.InsertCommand.Parameters("P_DESTINATION_TYPE_CODE").SourceColumn = "destination_type_code"
                    oda.InsertCommand.Parameters("P_INTERFACE_SOURCE_CODE").SourceColumn = "interface_source_code"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_FLAG").SourceColumn = "autosource_flag"
                    oda.InsertCommand.Parameters("P_DOCUMENT_TYPE_CODE").SourceColumn = "document_type_code"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_DOC_HEADER_ID").SourceColumn = "autosource_doc_header_id"
                    oda.InsertCommand.Parameters("P_AUTOSOURCE_DOC_LINE_NUM").SourceColumn = "autosource_doc_line_num"
                    oda.InsertCommand.Parameters("P_SUGGESTED_VENDOR_ID").SourceColumn = "suggested_vendor_id"
                    oda.InsertCommand.Parameters("P_SUGGESTED_VENDOR_SITE_ID").SourceColumn = "suggested_vendor_site_id"
                    oda.InsertCommand.Parameters("P_SUGGESTED_BUYER_ID").SourceColumn = "suggested_buyer_id"
                    oda.InsertCommand.Parameters("P_CREATED_BY").SourceColumn = "created_by"
                    oda.InsertCommand.Parameters("P_LAST_UPDATED_BY").SourceColumn = "last_updated_by"
                    oda.InsertCommand.Parameters("P_JUSTIFICATION").SourceColumn = "justification"
                    oda.InsertCommand.Connection.Open()
                    oda.Update(ds.Tables(0))
                    oda.InsertCommand.Connection.Close()
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step2 insert data to interface", "I")

                    oda.InsertCommand.Parameters.Clear()
                    oda.InsertCommand.CommandType = CommandType.StoredProcedure
                    oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.del_po_requisitions_inte"
                    oda.InsertCommand.Parameters.Add("o_return_message", OracleType.VarChar, 1000)
                    oda.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 10)
                    oda.InsertCommand.Parameters.Add("p_created_by", OracleType.Int32)
                    oda.InsertCommand.Parameters.Add("p_date_from", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_date_to", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_interface_source_code", OracleType.VarChar, 250)
                    oda.InsertCommand.Parameters("o_return_message").Direction = ParameterDirection.Output
                    oda.InsertCommand.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                    oda.InsertCommand.Parameters("p_created_by").Value = ConfigurationManager.AppSettings("AutoEJIT_User_id_ODay")
                    oda.InsertCommand.Parameters("p_date_from").Value = Format(DateTime.Now, "yyyy/MM/dd") & " 00:00:01"
                    oda.InsertCommand.Parameters("p_date_to").Value = Format(DateTime.Now, "yyyy/MM/dd") & " 23:59:59"
                    oda.InsertCommand.Parameters("p_interface_source_code").Value = "SKIPNEXTDATE"
                    oda.InsertCommand.Connection.Open()
                    oda.InsertCommand.ExecuteNonQuery()
                    oda.InsertCommand.Connection.Close()

                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_autoejit_save"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Parameters.Add("@o_flag", SqlDbType.NVarChar, 1).Direction = ParameterDirection.Output
                    sda.SelectCommand.Parameters.Add("@o_msg", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                    sda.SelectCommand.Connection.Open()
                    sda.SelectCommand.ExecuteNonQuery()
                    sda.SelectCommand.Connection.Close()
                    error_flag = sda.SelectCommand.Parameters("@o_flag").Value.ToString()
                    If error_flag <> "Y" Then
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "call ora_autoejit_save error: " & sda.SelectCommand.Parameters("@o_msg").Value.ToString(), "E")
                        Return False
                        Exit Function
                    End If
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step3 save upload data to EJIT table", "I")
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_get_autoejit_mo_orgcode"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(moorgcode)
                    sda.SelectCommand.Connection.Close()
                    For j = 0 To moorgcode.Tables(0).Rows.Count - 1
                        sda.SelectCommand.Parameters.Clear()
                        mods.Clear()
                        sda.SelectCommand.CommandType = CommandType.StoredProcedure
                        sda.SelectCommand.CommandText = "ora_get_autoejit_mo"
                        sda.SelectCommand.CommandTimeout = TimeOut_M5
                        sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 250).Value = moorgcode.Tables(0).Rows(j)("ORG_CODE")
                        sda.SelectCommand.Connection.Open()
                        sda.Fill(mods)
                        sda.SelectCommand.Connection.Close()
                        ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step4 get MO data", "I")
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.get_mo_batch_id"
                        oda.InsertCommand.Parameters.Add("o_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_batch_id").Direction = ParameterDirection.Output
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()

                        Batch_id = oda.InsertCommand.Parameters("o_batch_id").Value
                        For i = 0 To mods.Tables(0).Rows.Count - 1
                            If mods.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                                mods.Tables(0).Rows(i).SetAdded()
                            End If
                        Next
                        For i = 0 To mods.Tables(0).Rows.Count - 1
                            mods.Tables(0).Rows(i)("BATCH_ID") = Batch_id
                            mods.Tables(0).Rows(i)("USERID") = user_id
                        Next
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.insert_mo_data"
                        oda.InsertCommand.Parameters.Add("p_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                        oda.InsertCommand.Parameters.Add("p_mo_qty", OracleType.Double)
                        oda.InsertCommand.Parameters.Add("p_subinv", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_reference", OracleType.VarChar, 240)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                        oda.InsertCommand.Parameters("p_batch_id").SourceColumn = "batch_id"
                        oda.InsertCommand.Parameters("p_item_num").SourceColumn = "item_no"
                        oda.InsertCommand.Parameters("p_mo_qty").SourceColumn = "mo_qty"
                        oda.InsertCommand.Parameters("p_subinv").SourceColumn = "subinv"
                        oda.InsertCommand.Parameters("p_locator").SourceColumn = "locator"
                        oda.InsertCommand.Parameters("p_reference").SourceColumn = "reference"
                        oda.InsertCommand.Parameters("p_user_id").SourceColumn = "userid"
                        oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                        oda.InsertCommand.Parameters("o_msg").SourceColumn = "o_msg"
                        oda.InsertCommand.Connection.Open()
                        oda.Update(mods.Tables(0))
                        oda.InsertCommand.Connection.Close()
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.autoejit_mo_submit_req"
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 1000)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_batch_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_msg").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("p_org_code").Value = moorgcode.Tables(0).Rows(j)("ORG_CODE")
                        oda.InsertCommand.Parameters("p_batch_id").Value = Batch_id
                        oda.InsertCommand.Parameters("p_user_id").Value = user_id
                        If moorgcode.Tables(0).Rows(j)("ORG_CODE") = "680" Then
                            oda.InsertCommand.Parameters("p_resp_id").Value = wfoe_resp_id
                        Else
                            oda.InsertCommand.Parameters("p_resp_id").Value = resp_id
                        End If
                        oda.InsertCommand.Parameters("p_appl_id").Value = appl_id
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()

                        error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                        If error_flag <> "Y" Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "E")
                        End If
                        If Len(oda.InsertCommand.Parameters("o_msg").Value.ToString()) > 0 Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "I")
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step5 submit request to create MO", "I")
                    'Dim SourceCodeData As System.Data.DataTable
                    'Dim myDataColumn As DataColumn
                    'SourceCodeData = New Data.DataTable("InterfaceSourceCode")
                    'myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
                    'SourceCodeData.Columns.Add(myDataColumn)
                    'myDataColumn = New Data.DataColumn("SourceCode", System.Type.GetType("System.String"))
                    'SourceCodeData.Columns.Add(myDataColumn)
                    'ds.Tables.Add(SourceCodeData)
                    'scrow = ds.Tables("InterfaceSourceCode").NewRow
                    'scrow("SourceCode") = "AutoEJIT1530"
                    'scrow("OrgCode") = "580"
                    'scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_EJIT").ToString
                    'ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    'ds.Tables("InterfaceSourceCode").AcceptChanges()
                    'scrow = ds.Tables("InterfaceSourceCode").NewRow
                    'scrow("SourceCode") = "AUTOHUB1530"
                    'scrow("OrgCode") = "580"
                    'scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_HUB").ToString
                    'ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    'ds.Tables("InterfaceSourceCode").AcceptChanges()
                    'If ConfigurationManager.AppSettings("AutoEJIT_SC_IR").ToString <> "NO" Then
                    '    scrow = ds.Tables("InterfaceSourceCode").NewRow
                    '    scrow("OrgCode") = "580"
                    '    scrow("SourceCode") = ConfigurationManager.AppSettings("AutoEJIT_SC_IR").ToString
                    '    ds.Tables("InterfaceSourceCode").Rows.Add(scrow)
                    '    ds.Tables("InterfaceSourceCode").AcceptChanges()
                    ds.Tables(1).TableName = "InterfaceSourceCode"
                    For i = 0 To ds.Tables("InterfaceSourceCode").Rows.Count - 1
                        oda.InsertCommand.Parameters.Clear()
                        oda.InsertCommand.CommandType = CommandType.StoredProcedure
                        oda.InsertCommand.CommandText = "apps.xxetr_kanban_pkg.autoejit_po_submit_req"
                        oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 1000)
                        oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                        oda.InsertCommand.Parameters.Add("p_source_code", OracleType.VarChar, 100)
                        oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32)
                        oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32)
                        oda.InsertCommand.Parameters("o_msg").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                        oda.InsertCommand.Parameters("p_source_code").Value = ds.Tables("InterfaceSourceCode").Rows(i)("SourceCode")
                        oda.InsertCommand.Parameters("p_user_id").Value = user_id
                        If ds.Tables("InterfaceSourceCode").Rows(i)("OrgCode") = "680" Then
                            oda.InsertCommand.Parameters("p_resp_id").Value = wfoe_resp_id
                        Else
                            oda.InsertCommand.Parameters("p_resp_id").Value = resp_id
                        End If
                        oda.InsertCommand.Parameters("p_appl_id").Value = appl_id
                        oda.InsertCommand.Connection.Open()
                        oda.InsertCommand.ExecuteNonQuery()
                        oda.InsertCommand.Connection.Close()
                        error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                        If error_flag <> "Y" Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "E")
                        End If
                        If Len(oda.InsertCommand.Parameters("o_msg").Value.ToString()) > 0 Then
                            ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oda.InsertCommand.Parameters("o_msg").Value.ToString(), "I")
                        End If
                    Next
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step6 submit request to create PO/IR", "I")
                    Dim mailtype As Integer
                    Sqlstr = "select dbo.ora_autoejit_mailtype() mail_type"
                    mailtype = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                    ds.Tables.Clear()
                    Dim dt As New DataTable
                    Dim fileName As String
                    Dim fileNamepath As String
                    Dim sheetName As String
                    fileName = Format(DateTime.Now, "yyyyMMddHHmmss") & ".xlsx"
                    fileNamepath = "D:\AutoEJITErrorReport\" & fileName
                    sda.SelectCommand.Parameters.Clear()
                    sda.SelectCommand.CommandType = CommandType.StoredProcedure
                    sda.SelectCommand.CommandText = "ora_get_autoejit_errorlist"
                    sda.SelectCommand.CommandTimeout = TimeOut_M5
                    sda.SelectCommand.Connection.Open()
                    sda.Fill(ds)
                    sda.SelectCommand.Connection.Close()
                    sheetName = "Sheet1"
                    GenerateExcel(ds.Tables(0), fileNamepath, sheetName)
                    Dim mailsubject As String
                    Dim mailtext As String
                    Dim mailto() As String = {"charleszhao@artesyn.com", "janecao@artesyn.com", "MaggieMo@artesyn.com"}
                    Dim mailtouser() As String = {"charleszhao@artesyn.com", "janecao@artesyn.com", "MaggieMo@artesyn.com", ConfigurationManager.AppSettings("AutoEJIT_ErrMailGroup").ToString}
                    If mailtype = -1 Then
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram calculate error"
                        mailtext = "Failed"
                        SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext)
                    ElseIf mailtype = 0 Then
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram Error Message records list"
                        mailtext = "<font size='4' color='red'>Hi,ALL</font><br><font size='3' color='black'>"
                        mailtext = mailtext & "Please click the below path to check the error message for AutoEJIT program . Thanks</font>"
                        mailtext = mailtext & "<br><br><font size='4' color='blue'>\\CNAPGZHOFS10\AutoEJITErrorReport\" & fileName & "</font>"
                        SendMail("cnapgzhofs10@artesyn.com", mailtouser, mailsubject, mailtext)
                    Else
                        mailsubject = "FS10-" & Format(DateTime.Now, "yyyy-MM-dd") & " AutoEJITProgram run successfully"
                        mailtext = "Successful"
                        SendMail("cnapgzhofs10@artesyn.com", mailto, mailsubject, mailtext)
                    End If
                    ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", "Step7 Send mail via webservice", "I")
                End If
                Return True
            Catch oe As Exception
                ErrorLogging("AutoEJIT", "KB_SCHEDULER_OFF", oe.Message & oe.Source, "E")
                Return False
            Finally
                If l_count > 0 Then
                    If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                    If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
                End If
            End Try
        End Using
    End Function

    Public Function UploadAutoEJIT_DJSum_to_Oracle() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim error_flag As String
            Try
                oda = da.Oda_Insert_EJIT()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.delete_xxetr_bom_expand"
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 2000)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 150)
                oda.InsertCommand.Parameters("o_err_msg").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("p_identity").Value = "580-PREWORK"
                oda.InsertCommand.Connection.Open()
                oda.InsertCommand.ExecuteNonQuery()
                oda.InsertCommand.Connection.Close()
                error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                If error_flag <> "Y" Then
                    Return False
                    Exit Function
                End If
                oda.InsertCommand.Parameters.Clear()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.delete_xxetr_bom_expand"
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 2000)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 150)
                oda.InsertCommand.Parameters("o_err_msg").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("p_identity").Value = "680-PREWORK"
                oda.InsertCommand.Connection.Open()
                oda.InsertCommand.ExecuteNonQuery()
                oda.InsertCommand.Connection.Close()
                error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                If error_flag <> "Y" Then
                    Return False
                    Exit Function
                End If
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_autoejit_djsum"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@p_org_code", SqlDbType.VarChar, 150).Value = "580"
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                Dim i As Integer
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        ds.Tables(0).Rows(i).SetAdded()
                    End If
                Next
                oda.InsertCommand.Parameters.Clear()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.insert_bj_mo_data"
                oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 25)
                oda.InsertCommand.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.Double)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters("p_org_code").SourceColumn = "orgcode"
                oda.InsertCommand.Parameters("p_item_num").SourceColumn = "component"
                oda.InsertCommand.Parameters("p_locator").SourceColumn = "pwclocator"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "qty"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_msg").SourceColumn = "o_msg"
                oda.InsertCommand.Connection.Open()
                oda.Update(ds.Tables(0))
                oda.InsertCommand.Connection.Close()
                Dim dr() As DataRow = Nothing
                dr = ds.Tables(0).Select("o_flag = 'N'")
                If dr.Length > 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch oe As Exception
                ErrorLogging("UploadAutoEJIT_DJSum_to_Oracle", "eTrace_WS_Background", oe.Message & oe.Source, "E")
                Return False
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function UploadAutoEJIT_DJList_to_Oracle() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet
            Dim error_flag As String
            Try
                oda = da.Oda_Insert_EJIT()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.delete_xxetr_bom_expand"
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 2000)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 150)
                oda.InsertCommand.Parameters("o_err_msg").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("p_identity").Value = "580-PREWORK"
                oda.InsertCommand.Connection.Open()
                oda.InsertCommand.ExecuteNonQuery()
                oda.InsertCommand.Connection.Close()
                error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                If error_flag <> "Y" Then
                    Return False
                    Exit Function
                End If
                oda.InsertCommand.Parameters.Clear()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.delete_xxetr_bom_expand"
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 2000)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10)
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 150)
                oda.InsertCommand.Parameters("o_err_msg").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters("p_identity").Value = "680-PREWORK"
                oda.InsertCommand.Connection.Open()
                oda.InsertCommand.ExecuteNonQuery()
                oda.InsertCommand.Connection.Close()
                error_flag = oda.InsertCommand.Parameters("o_flag").Value.ToString()
                If error_flag <> "Y" Then
                    Return False
                    Exit Function
                End If
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_get_bj_mo"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@p_org_code", SqlDbType.VarChar, 150).Value = "580"
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                Dim i As Integer
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        ds.Tables(0).Rows(i).SetAdded()
                    End If
                Next
                oda.InsertCommand.Parameters.Clear()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.insert_bj_mo_data"
                oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 25)
                oda.InsertCommand.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.Double)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters("p_org_code").SourceColumn = "orgcode"
                oda.InsertCommand.Parameters("p_item_num").SourceColumn = "component"
                oda.InsertCommand.Parameters("p_locator").SourceColumn = "pwclocator"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "qty"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_msg").SourceColumn = "o_msg"
                oda.InsertCommand.Connection.Open()
                oda.Update(ds.Tables(0))
                oda.InsertCommand.Connection.Close()
                Dim dr() As DataRow = Nothing
                dr = ds.Tables(0).Select("o_flag = 'N'")
                If dr.Length > 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch oe As Exception
                ErrorLogging("UploadAutoEJIT_DJList_to_Oracle", "eTrace_WS_Background", oe.Message & oe.Source, "E")
                Return False
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function UploadBJData_to_Oracle() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet

            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_get_bj_mo"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Parameters.Add("@p_org_code", SqlDbType.VarChar, 150).Value = "580"
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()

                Dim i As Integer
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        ds.Tables(0).Rows(i).SetAdded()
                    End If
                Next

                oda = da.Oda_Insert_EJIT()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.insert_bj_mo_data"
                oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 25)
                oda.InsertCommand.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.Double)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_msg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput

                oda.InsertCommand.Parameters("p_org_code").SourceColumn = "orgcode"
                oda.InsertCommand.Parameters("p_item_num").SourceColumn = "component"
                oda.InsertCommand.Parameters("p_locator").SourceColumn = "pwclocator"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "qty"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_msg").SourceColumn = "o_msg"

                oda.InsertCommand.Connection.Open()
                oda.Update(ds.Tables(0))
                oda.InsertCommand.Connection.Close()

                Dim dr() As DataRow = Nothing
                dr = ds.Tables(0).Select("o_flag = 'N'")
                If dr.Length > 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch oe As Exception
                ErrorLogging("UploadBJData_to_Oracle", "eTrace_WS_Background", oe.Message & oe.Source, "E")
                Return False
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function LoginCheck(ByVal ERPLoginData As ERPLogin) As UserData
        Using da As DataAccess = GetDataAccess()
            LoginCheck = New UserData

            Try
                Dim UserType As String = ""
                Dim ConfigID As String = "HHPD"
                Dim Sqlstr, UserID, OSPCheck, ConfigVersion As String

                Dim Control As String
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'SYS001'")
                Control = da.ExecuteScalar(Sqlstr)
                If Control = 0 Then
                    LoginCheck.ErrorMsg = "Disable eTrace login now!"
                    Exit Function
                End If

                'Validate OSP User Login before Oracle User Login check ?
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'HHC001'")
                OSPCheck = Convert.ToString(da.ExecuteScalar(Sqlstr))
                If OSPCheck = "YES" Then
                    'First check if the User is OSP User, if yes, no need to verify account from Oracle
                    LoginCheck = Check_OSPUser(ERPLoginData)
                    If Not LoginCheck.User Is Nothing Then
                        If LoginCheck.ErrorMsg <> "" Then Exit Function

                        UserType = LoginCheck.UserType.ToUpper
                        If UserType = "OSP" Then ConfigID = "HHOSP"
                    End If
                End If

                'Get OrgID from T_Server
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)
                LoginCheck.OrgID = OrgID

                'Get the right version for HH: HHPD or HHOSP
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = '{0}'", ConfigID)
                ConfigVersion = Convert.ToString(da.ExecuteScalar(Sqlstr))
                LoginCheck.MinClientVersion = ConfigVersion


                ''Compare HH version and ConfigVersion to see if need to give error message
                'If (Not ERPLoginData.ClientVersion Is Nothing) AndAlso ConfigVersion <> "" Then
                '    Dim HHVer As Integer = Int(Replace(ERPLoginData.ClientVersion.ToString.Trim, ".", ""))
                '    Dim SqlVer As Integer = Int(Replace(ConfigVersion, ".", ""))
                '    If SqlVer > HHVer Then
                '        LoginCheck.ErrorMsg = "The Client Version is too low, please contact IT to upgrade first!"
                '        Exit Function
                '    End If
                'End If


                If UserType = "OSP" Then
                    Sqlstr = String.Format("Select tranid,module from t_transaction with (nolock) where Active = 1")
                    LoginCheck.TransactionID = da.ExecuteDataSet(Sqlstr, "TransactionID")
                    Exit Function
                End If

                'Validate Oracle account and get user information  if Control = 1
                If Control = 1 Then
                    If Not ERPLoginData.ClientVersion Is Nothing Then
                        LoginCheck = Get_UserData(ERPLoginData)
                    Else
                        LoginCheck = Get_UserInfo(ERPLoginData)
                    End If
                ElseIf Control = 2 Then
                    'No need to check Oracle account & password if Control = 2
                    LoginCheck.UserID = "Control_2"    'Warning: UserID can't set to ErpLoginData.User
                    LoginCheck.RespID_Inv = "Control_2"
                    LoginCheck.RespID_WIP = "Control_2"
                    LoginCheck.User = ERPLoginData.User.ToUpper
                    LoginCheck.FirstName = ERPLoginData.User.ToUpper
                    LoginCheck.LastName = ""
                    LoginCheck.Server = ""
                End If

                LoginCheck.OrgID = OrgID
                LoginCheck.MinClientVersion = ConfigVersion
                If LoginCheck.UserID Is Nothing Then Exit Function

                'Set user Type and Dept as blank for normal users
                LoginCheck.UserDept = ""
                LoginCheck.UserType = ""

                If ERPLoginData.Application = "HH" Then
                    'Codes below aded by Andyking for User Authorization Control ----------Begin  (Please don't delete!!!!!)
                    Sqlstr = String.Format("Select UserID from T_Users with (nolock) where UserID = '{0}'", LoginCheck.UserID)
                    UserID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If UserID = "" Then
                        Sqlstr = String.Format("Select tranid,module from t_transaction with (nolock) where Active = 1 and  not Module like '%SUPERUSER%'")
                        If Control = 2 Then
                            Sqlstr = String.Format("Select tranid,module from t_transaction with (nolock) where Active = 1 and  Module = 'ETRACE' ")
                        End If
                        LoginCheck.TransactionID = da.ExecuteDataSet(Sqlstr, "TransactionID")
                        Exit Function
                    End If

                    Sqlstr = String.Format("Select t_transaction.tranid,t_transaction.module from t_transaction inner join t_roleitem on t_transaction.tranid = t_roleitem.tranid inner join t_role on t_roleitem.roleid = t_role.roleid inner join t_userrole on t_role.roleid = t_userrole.roleid inner join t_users on t_userrole.userid = t_users.userid where t_transaction.Active = 1 AND t_users.userid = '{0}'", LoginCheck.UserID)
                    LoginCheck.TransactionID = da.ExecuteDataSet(Sqlstr, "TransactionID")
                    If LoginCheck.TransactionID.Tables.Count < 1 OrElse LoginCheck.TransactionID.Tables(0).Rows.Count < 1 Then
                        LoginCheck.ErrorMsg = "No eTrace authorization!"
                        Exit Function
                    End If

                    'Codes above aded by Andyking for User Authorization Control ----------End  (Please don't delete!!!!!)
                Else

                    'For PC Client program, needs to validate if the PC Module is ready for login
                    Dim TranID As String = ""
                    Sqlstr = String.Format("Select TranID from T_Transaction with (nolock) where TranID =  '{0}'  and Active = 'False' ", ERPLoginData.Application)
                    TranID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If TranID <> "" Then
                        LoginCheck.ErrorMsg = "System is locked now!"
                        Exit Function
                    End If
                End If

            Catch ex As Exception
                ErrorLogging("LoginCheck", LoginCheck.User, ex.Message & ex.Source, "E")
                LoginCheck.ErrorMsg = ex.Message.ToString
            End Try
        End Using

    End Function

    Public Function Get_UserInfo(ByVal ERPLoginData As ERPLogin) As UserData
        Using da As DataAccess = GetDataAccess()

            Get_UserInfo = New UserData
            Dim Oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("app_data")

                Oda.SelectCommand.CommandType = CommandType.StoredProcedure
                Oda.SelectCommand.CommandText = "apps.xxetr_hr_validatelogin_pkg.get_valid_user"

                Oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = ERPLoginData.OrgCode
                Oda.SelectCommand.Parameters.Add("p_application", OracleType.VarChar).Value = "HH"      'ERPLoginData.Application
                Oda.SelectCommand.Parameters.Add("p_user_name", OracleType.VarChar).Value = ERPLoginData.User.ToUpper()
                Oda.SelectCommand.Parameters.Add("p_password", OracleType.VarChar).Value = ERPLoginData.PWD
                Oda.SelectCommand.Parameters.Add("o_result", OracleType.VarChar, 4000).Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters.Add("o_app_data", OracleType.Cursor).Direction = ParameterDirection.Output

                If Oda.SelectCommand.Connection.State = ConnectionState.Closed Then
                    Oda.SelectCommand.Connection.Open()
                End If

                Oda.Fill(ds, "app_data")
                Oda.SelectCommand.Connection.Close()
                Get_UserInfo.Server = Oda.SelectCommand.Connection.DataSource.ToUpper

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    Get_UserInfo.ErrorMsg = "No authorization to transact "
                    Exit Function
                End If
                If ds.Tables(0).Columns.Count = 1 OrElse ds.Tables(0).Columns(0).ColumnName.ToUpper = "ERROR" Then
                    Get_UserInfo.ErrorMsg = ds.Tables(0).Rows(0)(0).ToString
                    Exit Function
                End If

                Dim UserInfo As String = Oda.SelectCommand.Parameters("o_result").Value.ToString
                If UserInfo = "" Then
                    Get_UserInfo.ErrorMsg = "No authorization to transact "
                    Exit Function
                End If

                Dim Arry() As String
                Arry = Split(UserInfo, ";")
                If Not Arry Is DBNull.Value Then
                    Get_UserInfo.User = Arry(0).ToString.ToUpper
                    Get_UserInfo.FirstName = Arry(3).ToString.ToUpper
                    Get_UserInfo.LastName = Arry(4).ToString.ToUpper
                    Get_UserInfo.UserID = Arry(6)
                End If

                Dim i As Integer
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i)("Type").ToString = "PO" Then
                        Get_UserInfo.AppID_PO = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserInfo.RespID_PO = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "INV" Then
                        Get_UserInfo.AppID_Inv = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserInfo.RespID_Inv = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "WIP" Then
                        Get_UserInfo.AppID_WIP = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserInfo.RespID_WIP = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "KB" Then
                        Get_UserInfo.AppID_KB = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserInfo.RespID_KB = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                Next

            Catch oe As Exception
                'Return Nothing
                Throw oe
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_UserData(ByVal ERPLoginData As ERPLogin) As UserData
        Using da As DataAccess = GetDataAccess()

            Get_UserData = New UserData
            Dim Oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("app_data")

                Oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'Oda.SelectCommand.CommandText = "apps.xxetr_hr_validatelogin_pkg.get_valid_user"
                Oda.SelectCommand.CommandText = "apps.xxetr_hr_validatelogin_pkg.get_valid_user_v2"

                Oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = ERPLoginData.OrgCode
                Oda.SelectCommand.Parameters.Add("p_application", OracleType.VarChar).Value = "HH"      'ERPLoginData.Application
                Oda.SelectCommand.Parameters.Add("p_user_name", OracleType.VarChar).Value = ERPLoginData.User.ToUpper()
                Oda.SelectCommand.Parameters.Add("p_password", OracleType.VarChar).Value = ERPLoginData.PWD
                Oda.SelectCommand.Parameters.Add("o_result", OracleType.VarChar, 4000).Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters.Add("o_app_data", OracleType.Cursor).Direction = ParameterDirection.Output

                If Oda.SelectCommand.Connection.State = ConnectionState.Closed Then
                    Oda.SelectCommand.Connection.Open()
                End If

                Oda.Fill(ds, "app_data")
                Oda.SelectCommand.Connection.Close()
                Get_UserData.Server = Oda.SelectCommand.Connection.DataSource.ToUpper

                If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                    Get_UserData.ErrorMsg = "No authorization to transact "
                    Exit Function
                End If

                If ds.Tables(0).Columns.Count = 1 OrElse ds.Tables(0).Columns(0).ColumnName.ToUpper = "ERROR" Then
                    Get_UserData.ErrorMsg = ds.Tables(0).Rows(0)(0).ToString
                    Exit Function
                End If

                Dim UserInfo As String = Oda.SelectCommand.Parameters("o_result").Value.ToString
                If UserInfo = "" Then
                    Get_UserData.ErrorMsg = "No authorization to transact "
                    Exit Function
                End If

                Dim Arry() As String
                Arry = Split(UserInfo, ";")
                If Not Arry Is DBNull.Value Then
                    Get_UserData.User = Arry(0).ToString.ToUpper
                    Get_UserData.FirstName = Arry(3).ToString.ToUpper
                    Get_UserData.LastName = Arry(4).ToString.ToUpper
                    Get_UserData.UserID = Arry(6)
                End If

                'Allow MIS User to login eTrace for view only
                If ds.Tables(0).Rows.Count = 0 Then
                    Dim Sqlstr As String
                    Dim UserID As String = ""
                    Sqlstr = String.Format("Select UserID from T_Users with (nolock) where Status = 'Active' and UserType = 'MIS' and UserID = '{0}'", ERPLoginData.User.ToUpper)
                    UserID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    If UserID = "" Then
                        Get_UserData.ErrorMsg = "No authorization to transact "
                        Exit Function
                    End If

                    Get_UserData.RespID_Inv = "MIS"
                    Get_UserData.RespID_WIP = "MIS"
                    Get_UserData.RespID_KB = "MIS"
                    Exit Function
                End If


                Dim i As Integer
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i)("Type").ToString = "PO" Then
                        Get_UserData.AppID_PO = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserData.RespID_PO = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "INV" Then
                        Get_UserData.AppID_Inv = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserData.RespID_Inv = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "WIP" Then
                        Get_UserData.AppID_WIP = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserData.RespID_WIP = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                    If ds.Tables(0).Rows(i)("Type").ToString = "KB" Then
                        Get_UserData.AppID_KB = ds.Tables(0).Rows(i)("Application_ID").ToString
                        Get_UserData.RespID_KB = ds.Tables(0).Rows(i)("Responsibility_ID").ToString
                    End If
                Next

            Catch oe As Exception
                'Return Nothing
                Throw oe
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function ChangePassword(ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_hr_validatelogin_pkg.change_password"
                OC.Parameters.Add("p_user_name", OracleType.VarChar).Value = LoginData.User.ToUpper
                OC.Parameters.Add("p_oldpassword", OracleType.VarChar, 20).Value = LoginData.PWD
                OC.Parameters.Add("p_newpassword", OracleType.VarChar, 50).Value = LoginData.Application  'Save as New Password
                OC.Parameters.Add("o_success_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Dim Flag, Msg As String
                Flag = OC.Parameters("o_success_flag").Value.ToString
                Msg = OC.Parameters("o_error_mssg").Value.ToString
                OC.Connection.Close()

                If Flag = "Y" Then        ' Y: Verification successful; N: Verification failed
                    ChangePassword = Flag
                Else
                    ChangePassword = Msg
                End If

            Catch oe As Exception
                ErrorLogging("PublicFunction-ChangePassword", LoginData.User, oe.Message & oe.Source, "E")
                ChangePassword = oe.Message
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using

    End Function

    Public Function ValidateItemRevision(ByVal OracleLoginData As ERPLogin, ByVal Item As String, ByVal Revision As String, ByVal MoveType As String) As ItemRevList
        Using da As DataAccess = GetDataAccess()
            Dim OC As OracleCommand = da.OraCommand()
            Try
                Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

                Dim ds As New DataSet()
                ds.Tables.Add("item_revlist")
                Dim mydatarow As DataRow
                Dim aa As Integer

                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.validate_revision"
                OC.Parameters.Add("p_item_num", OracleType.VarChar, 50).Value = Item.ToUpper()
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgID  'OracleLoginData.OrgCode
                OC.Parameters.Add("p_item_rev", OracleType.VarChar, 50).Value = Revision.ToUpper()
                OC.Parameters.Add("o_item_revlist", OracleType.LongVarChar, 3000).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_success_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then
                    OC.Connection.Open()
                End If
                aa = CInt(OC.ExecuteNonQuery())
                ValidateItemRevision.Flag = OC.Parameters("o_success_flag").Value.ToString
                ValidateItemRevision.Msg = OC.Parameters("o_error_mssg").Value.ToString
                ValidateItemRevision.RevList = OC.Parameters("o_item_revlist").Value.ToString
                OC.Connection.Close()
                'OC.Dispose()
                ' Value for Flag ----- E: Error; Y: Item valid and Revision controlled; N: Item valid but not Revision controlled
            Catch oe As Exception
                ErrorLogging(MoveType, OracleLoginData.User, oe.Message & oe.Source, "E")
                ValidateItemRevision.Flag = "E"
                ValidateItemRevision.Msg = oe.Message
                ValidateItemRevision.RevList = ""
                Throw oe
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function ValidateItemType(ByVal OracleLoginData As ERPLogin, ByVal Item As String, ByVal Revision As String, ByVal MoveType As String) As ItemType
        Using da As DataAccess = GetDataAccess()
            Dim OC As OracleCommand = da.OraCommand()
            Try
                Dim aa As Integer
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.get_itemtype"
                OC.Parameters.Add("p_item_num", OracleType.VarChar, 50).Value = Item.ToUpper()
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                OC.Parameters.Add("p_item_rev", OracleType.VarChar, 50).Value = Revision.ToUpper()
                OC.Parameters.Add("o_type_name", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_success_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then
                    OC.Connection.Open()
                End If
                aa = CInt(OC.ExecuteNonQuery())
                ValidateItemType.Type = OC.Parameters("o_type_name").Value.ToString
                ValidateItemType.Flag = OC.Parameters("o_success_flag").Value.ToString
                ValidateItemType.Msg = OC.Parameters("o_error_mssg").Value.ToString
                OC.Connection.Close()
                'OC.Dispose()
                'ValidateItemType.Flag -------- E: Error; Y: Item type in "FG,SA,RAW"; N: Item type not in "FG,SA,RAW"
            Catch oe As Exception
                ErrorLogging(MoveType, OracleLoginData.User, oe.Message & oe.Source, "E")
                ValidateItemType.Flag = "E"
                ValidateItemType.Msg = oe.Message
                ValidateItemType.Type = ""
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function ValidateSubLoc(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.valid_subinvloc"
                'OC.Parameters.Add("p_org_code", OracleType.VarChar).Value = LoginData.OrgCode
                OC.Parameters.Add("p_org_code", OracleType.VarChar).Value = OrgID
                OC.Parameters.Add("p_subinv", OracleType.VarChar, 20).Value = SubInv.ToUpper()
                OC.Parameters.Add("p_locator", OracleType.VarChar, 50).Value = Locator.ToUpper()
                OC.Parameters.Add("o_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then
                    OC.Connection.Open()
                End If
                aa = CInt(OC.ExecuteNonQuery())
                Dim Flag, Msg As String
                Flag = OC.Parameters("o_flag").Value.ToString
                Msg = OC.Parameters("o_msg").Value.ToString
                OC.Connection.Close()
                'OC.Dispose()

                If Flag = "E" Then        ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                    ValidateSubLoc = Msg
                Else
                    ValidateSubLoc = Flag
                End If
            Catch oe As Exception
                ErrorLogging("PublicFunction-ValidateSubLoc", LoginData.User, "SubInventory: " & SubInv & ", " & oe.Message & oe.Source, "E")
                ValidateSubLoc = oe.Message
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Get_LocatorSP(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()
            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.valid_subinvloc_sp"
                OC.Parameters.Add("p_org_code", OracleType.VarChar).Value = OrgID  'LoginData.OrgCode
                OC.Parameters.Add("p_subinv", OracleType.VarChar, 20).Value = SubInv.ToUpper()
                OC.Parameters.Add("p_locator", OracleType.VarChar, 50).Value = Locator.ToUpper()
                OC.Parameters.Add("o_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_sp", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Dim Flag, Msg, LocatorSP As String
                Flag = OC.Parameters("o_flag").Value.ToString
                Msg = OC.Parameters("o_msg").Value.ToString
                LocatorSP = OC.Parameters("o_sp").Value.ToString
                OC.Connection.Close()
                'OC.Dispose()

                If Flag = "E" Then        ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                    Get_LocatorSP = Msg
                Else
                    Get_LocatorSP = Flag & "," & LocatorSP
                End If


            Catch oe As Exception
                ErrorLogging("PublicFunction-Get_LocatorSP", LoginData.User, "SubInventory: " & SubInv & ", " & oe.Message & oe.Source, "E")
                Get_LocatorSP = oe.Message
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using

    End Function

    Public Function SlotCheck(ByVal LoginData As ERPLogin, ByRef myCLIDSlot As CLIDSlot) As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim Options As String
        Dim d1, d2, d3 As DateTime
        Try
            ErrorLogging("PublicFunction-SlotCheck", LoginData.User, "Slot Check already started", "I")

            myCommand = New SqlClient.SqlCommand("Select Value from T_Config WHERE ConfigID = 'HHF018'", myConn)
            myConn.Open()
            myCommand.CommandTimeout = TimeOut_M5
            objReader = myCommand.ExecuteReader()
            While objReader.Read()
                Options = objReader.GetValue(0)
            End While
            objReader.Close()
            myConn.Close()

            SlotCheck = ""
            Dim ds As New DataSet

            d1 = DateTime.Now()

            If Options = 0 Then
                Using da As DataAccess = GetDataAccess()
                    Try
                        Dim Sqlstr As String
                        Sqlstr = String.Format("exec sp_SlotCheck '{0}','{1}' ", myCLIDSlot.Slot, myCLIDSlot.OrgCode)
                        ds = da.ExecuteDataSet(Sqlstr, "Slot")

                        d2 = DateTime.Now()
                        ErrorLogging("SlotCheck", LoginData.User, "Option 0 - Start time: " & d1 & " Finishtime: " & d2, "I")

                        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                            SlotCheck = "Error while checking Slot"
                            Exit Function
                        End If

                        If ds.Tables(0).Rows.Count > 0 Then
                            myCLIDSlot.SlotType = ds.Tables(0).Rows(0)("SlotType").ToString
                            myCLIDSlot.SubInv = ds.Tables(0).Rows(0)("SubInv").ToString
                            myCLIDSlot.Locator = ds.Tables(0).Rows(0)("Locator").ToString
                            myCLIDSlot.SlotMsg = ds.Tables(0).Rows(0)("ErrMsg").ToString
                            SlotCheck = myCLIDSlot.SlotType

                            'Light on the LED Slot if it's available
                            If SlotCheck = "LA" OrElse SlotCheck = "LC" Then
                                Dim LightOn As Boolean = True
                                SlotLightOn(myCLIDSlot.Slot, LightOn, LoginData.User)
                            End If
                            If SlotCheck <> "S" Then Exit Function


                            'Invald Slot, then go to Oracle and check if it's Oracle SubInv or not
                            If myCLIDSlot.SlotType = "S" Then
                                myCLIDSlot.SlotMsg = ""

                                Dim CheckSLoc As String
                                Try
                                    CheckSLoc = ValidateSubLoc(LoginData, myCLIDSlot.Slot, "")
                                Catch ex As Exception
                                    ErrorLogging("SlotCheck-ValidateSubLoc", LoginData.User, ex.Message & ex.Source, "E")
                                    CheckSLoc = "Error while checking SubInv"
                                End Try

                                ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                                If CheckSLoc = "Y" OrElse CheckSLoc = "N" Then
                                    SlotCheck = CheckSLoc
                                Else
                                    SlotCheck = "E"
                                    myCLIDSlot.SlotMsg = CheckSLoc
                                End If
                                myCLIDSlot.SlotType = SlotCheck
                            End If

                        End If

                    Catch ex As Exception
                        ErrorLogging("PublicFunction-SlotCheck", LoginData.User, ex.Message & ex.Source, "E")
                        SlotCheck = "Error while checking Slot"
                    End Try
                End Using
            ElseIf Options = 1 Then
                Dim sda As SqlClient.SqlDataAdapter

                myConn.Open()
                myCommand = New SqlClient.SqlCommand("sp_SlotCheck", myConn)
                myCommand.CommandTimeout = TimeOut_M5
                myCommand.CommandType = CommandType.StoredProcedure
                myCommand.Parameters.AddWithValue("@Slot", myCLIDSlot.Slot)
                myCommand.Parameters.AddWithValue("@OrgCode", myCLIDSlot.OrgCode)
                sda = New SqlClient.SqlDataAdapter(myCommand)
                sda.Fill(ds)

                myConn.Close()

                d2 = DateTime.Now()
                ErrorLogging("SlotCheck", LoginData.User, "Option 1 - Start time: " & d1 & " Finishtime: " & d2, "I")

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    SlotCheck = "Error while checking Slot"
                    Exit Function
                End If

                If ds.Tables(0).Rows.Count > 0 Then
                    myCLIDSlot.SlotType = ds.Tables(0).Rows(0)("SlotType").ToString
                    myCLIDSlot.SubInv = ds.Tables(0).Rows(0)("SubInv").ToString
                    myCLIDSlot.Locator = ds.Tables(0).Rows(0)("Locator").ToString
                    myCLIDSlot.SlotMsg = ds.Tables(0).Rows(0)("ErrMsg").ToString
                    SlotCheck = myCLIDSlot.SlotType

                    'Light on the LED Slot if it's available
                    If SlotCheck = "LA" OrElse SlotCheck = "LC" Then
                        Dim LightOn As Boolean = True
                        SlotLightOn(myCLIDSlot.Slot, LightOn, LoginData.User)
                    End If
                    If SlotCheck <> "S" Then Exit Function


                    'Invald Slot, then go to Oracle and check if it's Oracle SubInv or not
                    If myCLIDSlot.SlotType = "S" Then
                        myCLIDSlot.SlotMsg = ""

                        Dim CheckSLoc As String
                        Try
                            CheckSLoc = ValidateSubLoc(LoginData, myCLIDSlot.Slot, "")
                        Catch ex As Exception
                            ErrorLogging("SlotCheck-ValidateSubLoc", LoginData.User, ex.Message & ex.Source, "E")
                            CheckSLoc = "Error while checking SubInv"
                        End Try

                        ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                        If CheckSLoc = "Y" OrElse CheckSLoc = "N" Then
                            SlotCheck = CheckSLoc
                        Else
                            SlotCheck = "E"
                            myCLIDSlot.SlotMsg = CheckSLoc
                        End If
                        myCLIDSlot.SlotType = SlotCheck
                    End If

                End If
            ElseIf Options = 2 Then
                Dim SubInv, Locator, SlotType, CLID, ErrMsg, CSFlag As String

                myCommand = New SqlClient.SqlCommand("Select SubInv, Locator, Remarks from T_InvSlot with (nolock) where Slot = @Slot and InvOrg = @OrgCode", myConn)
                myCommand.Parameters.Add("@Slot", SqlDbType.VarChar, 100, "Slot")
                myCommand.Parameters("@Slot").Value = myCLIDSlot.Slot
                myCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 100, "OrgCode")
                myCommand.Parameters("@OrgCode").Value = myCLIDSlot.OrgCode

                myConn.Open()
                myCommand.CommandTimeout = TimeOut_M5
                objReader = myCommand.ExecuteReader()
                While objReader.Read()
                    SubInv = objReader.GetValue(0)
                    Locator = objReader.GetValue(1)
                    CSFlag = objReader.GetValue(2)
                End While
                objReader.Close()
                myConn.Close()

                d2 = DateTime.Now()
                ErrorLogging("SlotCheck", LoginData.User, "Option 2 step1 - Start time: " & d1 & " Finishtime: " & d2, "I")

                If FixNull(SubInv) = "" Then
                    SlotType = "S"
                Else
                    myCommand = New SqlClient.SqlCommand("Select distinct CLID from T_CLMaster with (nolock) where StatusCode = '1' and StorageType = @Slot and OrgCode = @OrgCode", myConn)
                    myCommand.Parameters.Add("@Slot", SqlDbType.VarChar, 100, "Slot")
                    myCommand.Parameters("@Slot").Value = myCLIDSlot.Slot
                    myCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 100, "OrgCode")
                    myCommand.Parameters("@OrgCode").Value = myCLIDSlot.OrgCode

                    myConn.Open()
                    myCommand.CommandTimeout = TimeOut_M5
                    objReader = myCommand.ExecuteReader()
                    While objReader.Read()
                        CLID = objReader.GetValue(0)
                    End While
                    objReader.Close()
                    myConn.Close()
                    If FixNull(CLID) = "" Then
                        If CSFlag = "CS" Then
                            SlotType = "LC"                              'CS Slot Available           04/13/2019
                        Else
                            SlotType = "LA"                              'Normal Slot Available
                        End If
                    Else
                        SlotType = "LN"
                        ErrMsg = "Slot " & myCLIDSlot.Slot & " already linked to CLID: " & CLID
                    End If
                End If

                d3 = DateTime.Now()
                ErrorLogging("SlotCheck", LoginData.User, "Option 2 step2 - Start time: " & d2 & " Finishtime: " & d3, "I")

                myCLIDSlot.SlotType = SlotType
                myCLIDSlot.SubInv = SubInv
                myCLIDSlot.Locator = Locator
                myCLIDSlot.SlotMsg = ErrMsg
                SlotCheck = myCLIDSlot.SlotType

                'Light on the LED Slot if it's available
                If SlotCheck = "LA" OrElse SlotCheck = "LC" Then
                    ErrorLogging("SlotCheck", LoginData.User, "Option 2 step3: " & SlotCheck, "I")
                    Dim LightOn As Boolean = True
                    'SlotLightOn(myCLIDSlot.Slot, LightOn, LoginData.User)
                    ErrorLogging("SlotCheck", LoginData.User, "Option 2 step4: " & SlotCheck, "I")
                End If
                If SlotCheck <> "S" Then Exit Function


                'Invald Slot, then go to Oracle and check if it's Oracle SubInv or not
                If myCLIDSlot.SlotType = "S" Then
                    myCLIDSlot.SlotMsg = ""

                    Dim CheckSLoc As String
                    Try
                        CheckSLoc = ValidateSubLoc(LoginData, myCLIDSlot.Slot, "")
                    Catch ex As Exception
                        ErrorLogging("SlotCheck-ValidateSubLoc", LoginData.User, ex.Message & ex.Source, "E")
                        CheckSLoc = "Error while checking SubInv"
                    End Try

                    ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                    If CheckSLoc = "Y" OrElse CheckSLoc = "N" Then
                        SlotCheck = CheckSLoc
                    Else
                        SlotCheck = "E"
                        myCLIDSlot.SlotMsg = CheckSLoc
                    End If
                    myCLIDSlot.SlotType = SlotCheck
                End If
            ElseIf Options = 3 Then
                Dim TmpConfig, StrConfig(), StrCheck() As String
                Dim i, j As Integer

                If IsNumeric(myCLIDSlot.Slot) = True Then
                    Dim SubInv, Locator, SlotType, CLID, ErrMsg, CSFlag As String

                    myCommand = New SqlClient.SqlCommand("Select SubInv, Locator, Remarks from T_InvSlot with (nolock) where Slot = @Slot and InvOrg = @OrgCode", myConn)
                    myCommand.Parameters.Add("@Slot", SqlDbType.VarChar, 100, "Slot")
                    myCommand.Parameters("@Slot").Value = myCLIDSlot.Slot
                    myCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 100, "OrgCode")
                    myCommand.Parameters("@OrgCode").Value = myCLIDSlot.OrgCode

                    myConn.Open()
                    myCommand.CommandTimeout = TimeOut_M5
                    objReader = myCommand.ExecuteReader()
                    While objReader.Read()
                        SubInv = objReader.GetValue(0)
                        Locator = objReader.GetValue(1)
                        CSFlag = objReader.GetValue(2)
                    End While
                    objReader.Close()
                    myConn.Close()

                    Dim WSConfig As String = "WMS001"
                    If FixNull(SubInv) <> "" Then
                        SlotType = "LA"
                        If CSFlag = "CS" Then
                            SlotType = "LC"
                            WSConfig = "WMS006"
                        End If
                    End If


                    myCommand = New SqlClient.SqlCommand("Select Value from T_Config WHERE ConfigID =  @WSConfig ", myConn)
                    myCommand.Parameters.Add("@WSConfig", SqlDbType.VarChar, 100, "WSConfig")
                    myCommand.Parameters("@WSConfig").Value = WSConfig

                    myConn.Open()
                    myCommand.CommandTimeout = TimeOut_M5
                    objReader = myCommand.ExecuteReader()
                    While objReader.Read()
                        TmpConfig = objReader.GetValue(0)
                    End While
                    objReader.Close()
                    myConn.Close()
                    If FixNull(TmpConfig) <> "" Then
                        StrConfig = Split(TmpConfig, ";")
                        For i = LBound(StrConfig) To UBound(StrConfig)
                            If FixNull(StrConfig(i)) <> "" Then
                                Erase StrCheck
                                StrCheck = Split(StrConfig(i), ",")
                                For j = LBound(StrCheck) To UBound(StrCheck)
                                    If StrCheck(0) = myCLIDSlot.OrgCode Then
                                        If CSFlag = "CS" Then
                                            myCLIDSlot.SlotType = "LC"
                                        Else
                                            myCLIDSlot.SlotType = "LA"
                                        End If
                                        myCLIDSlot.SubInv = StrCheck(1)
                                        myCLIDSlot.Locator = StrCheck(2)
                                        myCLIDSlot.SlotMsg = ""
                                        SlotCheck = myCLIDSlot.SlotType

                                        Dim LightOn As Boolean = True
                                        SlotLightOn(myCLIDSlot.Slot, LightOn, LoginData.User)
                                        Exit Function
                                    End If
                                Next
                            End If
                        Next

                    End If
                Else
                    myCLIDSlot.SlotMsg = ""

                    Dim CheckSLoc As String
                    Try
                        CheckSLoc = ValidateSubLoc(LoginData, myCLIDSlot.Slot, "")
                    Catch ex As Exception
                        ErrorLogging("SlotCheck-ValidateSubLoc", LoginData.User, ex.Message & ex.Source, "E")
                        CheckSLoc = "Error while checking SubInv"
                    End Try

                    ' E: Error; Y: Valid and locator controlled; N: Valid but not locator controlled
                    If CheckSLoc = "Y" OrElse CheckSLoc = "N" Then
                        SlotCheck = CheckSLoc
                    Else
                        SlotCheck = "E"
                        myCLIDSlot.SlotMsg = CheckSLoc
                    End If
                    myCLIDSlot.SlotType = SlotCheck
                End If

            End If

            ErrorLogging("PublicFunction-SlotCheck", LoginData.User, "Slot Check already finished", "I")


        Catch ex As Exception
            ErrorLogging("SlotCheck", LoginData.User, ex.Message & ex.Source, "E")
            Exit Function
        End Try
    End Function

    Public Function SlotLightOn(ByVal SlotLists As String, ByVal LightOn As Boolean, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim dsSlot As New DataSet("DS")

            Dim myDR As DataRow
            Dim dtSlot As DataTable
            dtSlot = New DataTable("dtSlot")
            dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
            dsSlot.Tables.Add(dtSlot)

            Try
                Dim i As Integer
                Dim Arry() As String
                Arry = Split(SlotLists, ",")
                For i = 0 To Arry.Length - 1
                    Dim mySlot As String = Arry(i).ToString.Trim
                    If mySlot <> "" Then
                        myDR = dtSlot.NewRow()
                        myDR("Slot") = mySlot
                        dtSlot.Rows.Add(myDR)
                    End If
                Next

                '==Code--0/1/2(Off/On/blink. data type: integer)
                '==Interval--time(data type: integer. 0:not limit)
                Dim Code As Integer = 0
                If LightOn = True Then Code = 1

                Dim myWMS As WMS = New WMS
                'SlotLightOn = myWMS.LEDControlBySlot(dsSlot, Code, 3)
                SlotLightOn = myWMS.LEDControlBySlot(dsSlot, Code, 5)

            Catch ex As Exception
                ErrorLogging("PublicFunction-SlotLightOn", User, ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using
    End Function

    Public Function DockSlotLightOn(ByVal SlotLists As String, ByVal LightOn As Boolean, ByVal User As String, ByVal Interval As Integer) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim dsSlot As New DataSet("DS")

            Dim myDR As DataRow
            Dim dtSlot As DataTable
            dtSlot = New DataTable("dtSlot")
            dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
            dsSlot.Tables.Add(dtSlot)

            Try
                Dim i As Integer
                Dim Arry() As String
                Arry = Split(SlotLists, ",")
                For i = 0 To Arry.Length - 1
                    Dim mySlot As String = Arry(i).ToString.Trim
                    If mySlot <> "" Then
                        myDR = dtSlot.NewRow()
                        myDR("Slot") = mySlot
                        dtSlot.Rows.Add(myDR)
                    End If
                Next

                '==Code--0/1/2(Off/On/blink. data type: integer)
                '==Interval--time(data type: integer. 0:not limit)
                Dim Code As Integer = 0
                If LightOn = True Then Code = 1

                Dim myWMS As WMS = New WMS
                'SlotLightOn = myWMS.LEDControlBySlot(dsSlot, Code, 3)
                DockSlotLightOn = myWMS.LEDControlBySlot(dsSlot, Code, Interval)

            Catch ex As Exception
                ErrorLogging("PublicFunction-DockSlotLightOn", User, ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using
    End Function

    Public Function Get_ReasonCode(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("Reason")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_reason"

                oda.SelectCommand.Parameters.Add("o_reason", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "Reason")
                oda.SelectCommand.Connection.Close()

                ' ErrorLogging("Get_Reason", "Yong", oda.SelectCommand.Parameters("o_reason").Value.ToString.ToUpper)

                Return ds
            Catch oe As Exception
                ErrorLogging("Get_ReasonCode", OracleLoginData.User, oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_TransactionTypes(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("TransactionTypes")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_transaction_types"

                oda.SelectCommand.Parameters.Add("o_tran_data", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "TransactionTypes")
                oda.SelectCommand.Connection.Close()

                ' ErrorLogging("Get_Reason", "Yong", oda.SelectCommand.Parameters("o_reason").Value.ToString.ToUpper)

                Return ds
            Catch oe As Exception
                ErrorLogging("Get_TransactionTypes", OracleLoginData.User, oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_SubinvLoc(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

                Dim ds As New DataSet()
                ds.Tables.Add("SubinvLoc")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_org_sub_loc"

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = OrgID  'OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_sub_loc_data", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "SubinvLoc")
                oda.SelectCommand.Connection.Close()

                Return ds
            Catch oe As Exception
                ErrorLogging("Get_SubinvLoc", OracleLoginData.User, oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function get_iteminfo(ByVal ItemList As DataSet, ByVal oraclelogindata As ERPLogin, ByVal movetype As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = oraclelogindata.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = oraclelogindata.AppID_Inv    '401
        Dim flag As Boolean
        Try
            comm.CommandType = CommandType.StoredProcedure
            'comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            'comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(oraclelogindata.UserID)  ''15784
            'comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            'comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            'comm.ExecuteOracleNonQuery(aa)
            'comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_wip_pkg.get_item_master_a"   '53330   401
            comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
            comm.Parameters.Add("o_item_desc", OracleType.VarChar, 240)
            comm.Parameters.Add("o_commdity_code", OracleType.VarChar, 240)
            comm.Parameters.Add("o_item_rev", OracleType.VarChar, 240)
            comm.Parameters.Add("o_type_name", OracleType.VarChar, 240)
            comm.Parameters.Add("o_uom_code", OracleType.VarChar, 240)
            comm.Parameters.Add("o_routing_id", OracleType.Int32)
            comm.Parameters.Add("o_shelf_life_days", OracleType.Int32)
            comm.Parameters.Add("o_revision_control_code", OracleType.VarChar, 240)
            comm.Parameters.Add("o_lot_control_code", OracleType.VarChar, 240)
            comm.Parameters.Add("o_item_msl", OracleType.VarChar, 240)
            comm.Parameters.Add("o_item_temp_ctrl", OracleType.VarChar, 240)
            comm.Parameters.Add("o_valid_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_revlist", OracleType.VarChar, 240)
            comm.Parameters.Add("o_sublist", OracleType.LongVarChar, 5000)
            comm.Parameters.Add("o_shelf_life_ctrl", OracleType.Int32)
            comm.Parameters.Add("o_rohs_class", OracleType.VarChar, 50)

            comm.Parameters("o_item_desc").Direction = ParameterDirection.Output
            comm.Parameters("o_commdity_code").Direction = ParameterDirection.Output
            comm.Parameters("o_item_rev").Direction = ParameterDirection.Output
            comm.Parameters("o_type_name").Direction = ParameterDirection.Output
            comm.Parameters("o_uom_code").Direction = ParameterDirection.Output
            comm.Parameters("o_routing_id").Direction = ParameterDirection.Output
            comm.Parameters("o_shelf_life_days").Direction = ParameterDirection.Output
            comm.Parameters("o_revision_control_code").Direction = ParameterDirection.Output
            comm.Parameters("o_lot_control_code").Direction = ParameterDirection.Output
            comm.Parameters("o_item_msl").Direction = ParameterDirection.Output
            comm.Parameters("o_item_temp_ctrl").Direction = ParameterDirection.Output
            comm.Parameters("o_valid_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_revlist").Direction = ParameterDirection.Output
            comm.Parameters("o_sublist").Direction = ParameterDirection.Output
            comm.Parameters("o_shelf_life_ctrl").Direction = ParameterDirection.Output
            comm.Parameters("o_rohs_class").Direction = ParameterDirection.Output

            comm.Parameters("p_item_num").SourceColumn = "p_item_num"
            comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("o_item_desc").SourceColumn = "o_item_desc"
            comm.Parameters("o_commdity_code").SourceColumn = "o_commdity_code"
            comm.Parameters("o_item_rev").SourceColumn = "o_item_rev"
            comm.Parameters("o_type_name").SourceColumn = "o_type_name"
            comm.Parameters("o_uom_code").SourceColumn = "o_uom_code"
            comm.Parameters("o_routing_id").SourceColumn = "o_routing_id"
            comm.Parameters("o_shelf_life_days").SourceColumn = "o_shelf_life_days"
            comm.Parameters("o_revision_control_code").SourceColumn = "o_revision_control_code"
            comm.Parameters("o_lot_control_code").SourceColumn = "o_lot_control_code"
            comm.Parameters("o_item_msl").SourceColumn = "o_item_msl"
            comm.Parameters("o_item_temp_ctrl").SourceColumn = "o_item_temp_ctrl"
            comm.Parameters("o_valid_flag").SourceColumn = "o_valid_flag"
            comm.Parameters("o_revlist").SourceColumn = "o_revlist"
            comm.Parameters("o_sublist").SourceColumn = "o_sublist"
            comm.Parameters("o_shelf_life_ctrl").SourceColumn = "o_shelf_life_ctrl"
            comm.Parameters("o_rohs_class").SourceColumn = "o_rohs_class"

            oda_h.InsertCommand = comm
            oda_h.Update(ItemList.Tables("ItemData"))

            Return ItemList
        Catch ex As Exception
            ErrorLogging(movetype & "get_iteminfo", oraclelogindata.User, ex.Message & ex.Source, "E")
            If comm.Connection.State = ConnectionState.Open Then
                comm.Transaction.Rollback()
                comm.Connection.Close()
            End If
            Return ItemList
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function get_itemonhand(ByVal ItemList As DataSet, ByVal oraclelogindata As ERPLogin, ByVal movetype As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = oraclelogindata.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = oraclelogindata.AppID_Inv    '401
        Dim flag As Boolean
        Try
            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(oraclelogindata.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_wip_pkg.get_item_ava_qty"   '53330   401
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_rev", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("o_available_qty", OracleType.Double)
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 500)

            comm.Parameters("o_available_qty").Direction = ParameterDirection.Output
            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("p_item_num").SourceColumn = "p_item_num"
            comm.Parameters("p_item_rev").SourceColumn = "p_item_rev"
            comm.Parameters("p_subinventory").SourceColumn = "p_subinventory"
            comm.Parameters("p_locator").SourceColumn = "p_locator"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("o_available_qty").SourceColumn = "o_available_qty"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(ItemList.Tables("ItemData"))

            Return ItemList
        Catch ex As Exception
            ErrorLogging(movetype & "get_itemonhand", oraclelogindata.User, ex.Message & ex.Source, "E")
            If comm.Connection.State = ConnectionState.Open Then
                comm.Transaction.Rollback()
                comm.Connection.Close()
            End If
            Return ItemList
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function Get_AccountAlias(ByVal OracleLoginData As ERPLogin, ByVal Type As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("AccountAlias")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_account_alias"

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_acc_data", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "AccountAlias")
                oda.SelectCommand.Connection.Close()

                ' ErrorLogging("Get_Reason", "Yong", oda.SelectCommand.Parameters("o_reason").Value.ToString.ToUpper)
                If Type = "receipt" Or Type = "MiscRecpt" Then
                    Dim DR() As DataRow = Nothing
                    Dim i As Integer
                    DR = ds.Tables("AccountAlias").Select("accountalias like '%ISSUE%'")
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    'da.ExecuteNonQuery(String.Format("delete ds.tables(0) where accountalias like '{0}'", "%ISSUE%"))
                ElseIf Type = "issue" Then
                    Dim DR() As DataRow = Nothing
                    Dim i As Integer
                    DR = ds.Tables("AccountAlias").Select("accountalias like '%RECPT%' or accountalias like '%RCPT%' or accountalias like '%RECEIPT%'")
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    'da.ExecuteNonQuery(String.Format("delete ds.tables(0) where accountalias like '{0}' or accountalias like '{1}'", "%RECPT%", "%RECEIPT%"))
                Else

                End If
                ds.AcceptChanges()
                Return ds
            Catch oe As Exception
                ErrorLogging("Get_AccountAlias", OracleLoginData.User, oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetSOLine(ByVal LoginData As ERPLogin, ByVal SONo As String, ByVal SOLine As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim saleline() As String = Split(SOLine, ".")
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("SOData")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_oe_pkg.get_so_item"
                oda.SelectCommand.Parameters.Add("p_order_num", OracleType.VarChar).Value = SONo
                If saleline.Length = 2 Then
                    oda.SelectCommand.Parameters.Add("p_line_num", OracleType.VarChar).Value = saleline(0)
                    oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.VarChar).Value = saleline(1)

                Else
                    oda.SelectCommand.Parameters.Add("p_line_num", OracleType.VarChar).Value = ""
                    oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.VarChar).Value = ""
                End If

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "SOData")
                oda.SelectCommand.Connection.Close()

                Return ds

            Catch oe As Exception
                ErrorLogging("PublicFunction-GetSOLine", LoginData.User, "SONo/SOLine: " & SONo & "/" & SOLine & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Get_Subinv_Restrict(ByVal OracleLoginData As ERPLogin, ByVal TransType As String, ByVal AcctAliasName As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num

            Try
                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_subinv_restriction"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("p_tran_type", OracleType.VarChar, 300).Value = TransType
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_acc_name", OracleType.VarChar, 200).Value = AcctAliasName
                oda.SelectCommand.Parameters.Add("o_res_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_sub_from", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_sub_to", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_reason_data", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 4 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "subinv_rest"
                ds.Tables(1).TableName = "from_subinv"
                ds.Tables(2).TableName = "to_subinv"
                ds.Tables(3).TableName = "reason_code"
                Return ds

            Catch oe As Exception
                ErrorLogging("Get_SubInv_Restrict", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetCOOLists() As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_coo_data"
                oda.SelectCommand.Parameters.Add("o_coo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "COOData"
                Return ds

            Catch oe As Exception
                ErrorLogging("PublicFunction-GetCOOLists", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Post_SubInvTransfer(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv
        Dim appl As Integer = OracleLoginData.AppID_Inv

        Dim DJ As String = p_ds.Tables("transfer_table").Rows(0)("p_transaction_reference").ToString
        Dim ErrorMsg As String = "SubmitID " & TransactionID & " for DJ " & DJ

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.subinventory_transfer_batch"
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240)  ' p_transaction_type_name
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("p_subinventory_destination").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
            comm.Parameters("p_locator_source").SourceColumn = "p_locator_source"
            comm.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
            comm.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("p_transaction_type_name").SourceColumn = "p_transaction_type_name"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("transfer_table"))

            Dim i As Integer
            Dim DR() As DataRow = Nothing

            'DR = p_ds.Tables("transfer_table").Select("o_return_status = 'N'")
            DR = p_ds.Tables("transfer_table").Select("o_return_status = 'N' or o_return_status = ' ' or o_return_status IS Null ")

            If DR.Length = 0 Then
                'Record the transfer row count for submit
                'Dim ErrMsg As String = ErrorMsg & " with transfer Row Count: " & p_ds.Tables("transfer_table").Rows.Count.ToString
                'ErrorLogging("Post_SubInvTransfer-" & MoveType, OracleLoginData.User, ErrMsg, "I")

                Dim PostTransfer As String
                Dim PickDJ As String = MoveType & "/DJ: " & DJ

                PostTransfer = Submit_Transfer(OracleLoginData, TransactionID, PickDJ)
                If PostTransfer <> "Y" Then
                    If PostTransfer = "" Then PostTransfer = "Post failed"
                    p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("transfer_table").Rows(0)("o_return_message") = PostTransfer

                    Dim SubmitMsg As String
                    SubmitMsg = ErrorMsg & " , " & PostTransfer
                    ErrorLogging("Post_SubInvTransfer2-" & MoveType, OracleLoginData.User, SubmitMsg, "I")
                    del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
                End If
            Else

                For i = 0 To DR.Length - 1
                    Dim ErrMsg As String = ErrorMsg & " and flag: " & DR(i)("o_return_status").ToString & " and error message; " & DR(i)("o_return_message").ToString
                    ErrorLogging("Post_SubInvTransfer1-" & MoveType, OracleLoginData.User, ErrMsg, "I")
                Next
                del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
            End If

            comm.Connection.Close()
            Return p_ds

        Catch ex As Exception
            p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging("Post_SubInvTransfer-" & MoveType, OracleLoginData.User, ErrorMsg & ", " & ex.Message & ex.Source, "E")
            del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function Submit_Transfer(ByVal LoginData As ERPLogin, ByVal TransactionID As Integer, ByVal MoveType As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As OracleString
            Dim Oda As OracleCommand = da.OraCommand
            Dim transaction As OracleTransaction
            Dim bb As Integer
            Dim MsgFlag, ErrMsg As String
            Dim SubmitMsg As String

            Dim ModuleName As String = "Submit_Transfer-" & MoveType
            Dim ErrorMsg As String = "SubmitID " & TransactionID & " for " & MoveType

            Try
                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If
                transaction = Oda.Connection.BeginTransaction(IsolationLevel.ReadCommitted)
                Oda.Transaction = transaction

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.mtl_process_online"
                Oda.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_error_mssg", OracleType.VarChar, 10000)
                Oda.Parameters("o_success_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_error_mssg").Direction = ParameterDirection.Output

                bb = CInt(Oda.ExecuteOracleNonQuery(aa))
                MsgFlag = FixNull(Oda.Parameters("o_success_flag").Value)
                ErrMsg = FixNull(Oda.Parameters("o_error_mssg").Value)
                If MsgFlag = "Y" Then
                    SubmitMsg = MsgFlag
                    transaction.Commit()
                Else
                    SubmitMsg = "submit flag: " & MsgFlag & "; and error message: " & ErrMsg
                    transaction.Rollback()
                End If
                Oda.Connection.Close()
                Return DirectCast(SubmitMsg, String)

            Catch oe As Exception
                ErrorLogging(ModuleName, LoginData.User.ToUpper, ErrorMsg & ", " & oe.Message & oe.Source, "E")
                del_transfer_inte(LoginData.UserID, TransactionID)
                Return "Submit failed"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function del_transfer_inte(ByVal p_user_id As Int32, ByVal p_tran_head_id As Integer) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()
            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.del_mtl_tran_inte"
                Oda.Parameters.Add("p_tran_head_id", OracleType.Int32).Value = p_tran_head_id
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = p_user_id
                Oda.Parameters.Add("o_succ_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_return_message", OracleType.VarChar, 240)
                Oda.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_return_message").Direction = ParameterDirection.Output

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If
                aa = CInt(Oda.ExecuteNonQuery())
                MsgFlag = Oda.Parameters("o_succ_flag").Value
                Oda.Connection.Close()
                Return DirectCast(MsgFlag, String)

            Catch oe As Exception
                ErrorLogging("PublicFunction-del_transfer_inte", "", "tran_head_id: " & p_tran_head_id & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

#End Region

#Region "PrintLabel-SIT-Test"

    Public Function SIT_MassPrintCLIDs(ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs As DataSet = New DataSet
            SIT_MassPrintCLIDs = False

            Try
                Dim Sqlstr As String

                Sqlstr = String.Format("Select CLID from tmp_CLID ")
                myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                ElseIf myCLIDs.Tables(0).Rows.Count > 0 Then

                    SIT_MassPrintCLIDs = PrintCLIDs(myCLIDs, Printer)
                End If


            Catch ex As Exception
                ErrorLogging("SIT_MassPrintCLIDs", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using

    End Function

    Public Function SIT_DeleteCLIDs() As Boolean
        Using da As DataAccess = GetDataAccess()
            SIT_DeleteCLIDs = False

            Try
                Dim Sqlstr As String

                Sqlstr = String.Format("DELETE from tmp_CLID ")
                da.ExecuteNonQuery(Sqlstr)

                SIT_DeleteCLIDs = True

            Catch ex As Exception
                ErrorLogging("SIT_DeleteCLIDs", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using

    End Function

    Public Function SIT_Print_OnHand(ByVal Printer As String) As Boolean

        Using da As DataAccess = GetDataAccess()

            Dim ds As DataSet = New DataSet
            Dim myCLIDs As DataSet

            Try
                Dim Sqlstr As String

                Sqlstr = String.Format("Select OrgCode,MaterialNo from tmp_MaterialNo ")
                ds = da.ExecuteDataSet(Sqlstr, "Materials")


                Sqlstr = String.Format("DELETE from tmp_OnHand ")
                da.ExecuteNonQuery(Sqlstr)

                Dim i As Integer
                Dim OrgCode, MaterialNo As String

                If ds.Tables(0).Rows.Count > 0 Then

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        OrgCode = ds.Tables(0).Rows(i)(0)
                        MaterialNo = ds.Tables(0).Rows(i)(1)

                        'OrgCode,MaterialNo,CLID,QtyBaseUOM,BaseUOM,RecDocNo,RecDate,ExpDate,RTLot,SLOC,StorageBin,StockType,Manufacturer,ManufacturerPN,QMLStatus

                        myCLIDs = New DataSet
                        Sqlstr = String.Format("Select OrgCode,MaterialNo,CLID,QtyBaseUOM,BaseUOM,RecDocNo,RTLot,SLOC,StorageBin,StockType,Manufacturer,ManufacturerPN,QMLStatus from T_CLMaster where StatusCode = 1 and Not SLOC is NULL and Not StorageBin is NULL and MaterialNo = '{0}' and OrgCode = '{1}'", MaterialNo, OrgCode)
                        myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                        If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                        ElseIf myCLIDs.Tables(0).Rows.Count > 0 Then
                            SaveCLID(myCLIDs)
                        End If
                    Next


                    Sqlstr = String.Format("Select CLID from tmp_OnHand where StorageBin <> ' ' ")
                    myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                    If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                    ElseIf myCLIDs.Tables(0).Rows.Count > 0 Then

                        SIT_Print_OnHand = PrintCLIDs(myCLIDs, Printer)
                    End If


                End If

            Catch ex As Exception
                ErrorLogging("SIT_Print_OnHand", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using

    End Function

    Public Function SaveCLID(ByVal Items As DataSet) As Boolean
        Using da As DataAccess = GetDataAccess()
            SaveCLID = False

            Try
                Dim i As Integer
                Dim sqlstr As String

                For i = 0 To Items.Tables(0).Rows.Count - 1
                    Dim OrgCode, MaterialNo, CLID, BaseUOM, RecDocNo, RTLot, SLOC, StorageBin, StockType, Manufacturer, ManufacturerPN, QMLStatus As String
                    Dim QtyBaseUOM As Double

                    OrgCode = Items.Tables(0).Rows(i)("OrgCode").ToString
                    MaterialNo = Items.Tables(0).Rows(i)("MaterialNo").ToString
                    CLID = Items.Tables(0).Rows(i)("CLID").ToString
                    BaseUOM = Items.Tables(0).Rows(i)("BaseUOM").ToString

                    RecDocNo = Items.Tables(0).Rows(i)("RecDocNo").ToString
                    RTLot = Items.Tables(0).Rows(i)("RTLot").ToString

                    SLOC = Items.Tables(0).Rows(i)("SLOC").ToString
                    StorageBin = Items.Tables(0).Rows(i)("StorageBin").ToString
                    StockType = Items.Tables(0).Rows(i)("StockType").ToString
                    Manufacturer = Items.Tables(0).Rows(i)("Manufacturer").ToString
                    ManufacturerPN = Items.Tables(0).Rows(i)("ManufacturerPN").ToString
                    QMLStatus = Items.Tables(0).Rows(i)("QMLStatus").ToString

                    'RecDate = IIf(Items.Tables(0).Rows(i)("RecDate") Is DBNull.Value, DBNull.Value, CDate(Items.Tables(0).Rows(i)("RecDate")))
                    'ExpDate = IIf(Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value, DBNull.Value, CDate(Items.Tables(0).Rows(i)("ExpDate")))


                    QtyBaseUOM = IIf(Items.Tables(0).Rows(i)("QtyBaseUOM") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("QtyBaseUOM"))

                    sqlstr = String.Format("INSERT INTO tmp_CLID (OrgCode,MaterialNo,CLID,QtyBaseUOM,BaseUOM,RecDocNo,RTLot,SLOC,StorageBin,StockType,Manufacturer,ManufacturerPN,QMLStatus) values ('{0}','{1}','{2}','{3}', '{4}','{5}','{6}','{7}', '{8}','{9}','{10}','{11}', '{12}')", OrgCode, MaterialNo, CLID, QtyBaseUOM, BaseUOM, RecDocNo, RTLot, SLOC, StorageBin, StockType, Manufacturer, ManufacturerPN, QMLStatus)
                    da.ExecuteNonQuery(sqlstr)

                Next
                SaveCLID = True

            Catch ex As Exception
                ErrorLogging("SaveCLID", "", ex.Message & ex.Source, "E")
                SaveCLID = False
            End Try

        End Using


    End Function

#End Region

    Public Function FilterSpecial(ByVal SpecStr As String) As String
        If SpecStr = "" Then Return ""

        SpecStr = SpecStr.Replace("'", "''")
        FilterSpecial = SpecStr

    End Function

    Function SQLString(ByVal str As String) As String
        Dim pos As Integer
        pos = InStr(str, "'")
        While pos > 0
            str = Mid(str, 1, pos) & "'" & Mid(str, pos + 1)
            pos = InStr(pos + 2, str, "'")
        End While
        SQLString = str
    End Function

    Public Function FixNull(ByVal vMayBeNull As Object) As String
        On Error Resume Next
        FixNull = vbNullString & vMayBeNull
    End Function

    'Dataset transfer to XML for storeprocedure calling
    Public Function DStoXML(ByVal DS As DataSet) As String
        Return DS.GetXml()
    End Function

    Public Function DTtoXML(ByVal dt As DataTable, Optional ByVal dsName As String = "") As String

        Dim result As String = String.Empty
        Using sw As New IO.StringWriter()
            dt.WriteXml(sw)
            result = sw.ToString()
        End Using

        Return result


    End Function

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Public Function OpenPrintMatlLabel() As DataSet
        OpenPrintMatlLabel = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintMatlLabel.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintMatlLabel.Tables.Add(Contents)

        Dim cmdReadHeader, cmdReadItem As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header
        Dim getLabel As String
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
            & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
            & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
            & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
            & " FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1' and LabelFile ='CLMatLabel.lab' GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
            & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 10 ORDER BY CreatedOn"
        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1'", myConn)


        Try
            myConn.Open()
            drHeader = cmdReadHeader.ExecuteReader()
            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintMatlLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function OpenPrintNoMatlLabel() As DataSet
        OpenPrintNoMatlLabel = New DataSet

        Dim myDataColumn As DataColumn
        Dim mydatarow As DataRow

        Dim Labels As DataTable
        Labels = New Data.DataTable("Labels")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelFile", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LabelPrinter", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Content", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PrintedOn", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        Labels.Columns.Add(myDataColumn)
        OpenPrintNoMatlLabel.Tables.Add(Labels)

        Dim Contents As DataTable
        Contents = New Data.DataTable("Contents")
        myDataColumn = New Data.DataColumn("LabelSeqNo", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("title", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("detail", System.Type.GetType("System.String"))
        Contents.Columns.Add(myDataColumn)
        OpenPrintNoMatlLabel.Tables.Add(Contents)

        Dim cmdReadHeader, cmdReadItem As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim drHeader As SqlClient.SqlDataReader
        Dim arryDetail() As String
        Dim i As Integer
        'Read from PDTO Header
        Dim getLabel As String
        'getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
        '         & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
        '         & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.LabelSeqNo) AS 'n'" _
        '         & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
        '         & " FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1' and LabelFile <> 'CLMatLabel_tmp.lab' GROUP BY LabelPrinter, LabelSeqNo, LabelFile, [Content]," _
        '         & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 10 ORDER BY LabelSeqNo"
        getLabel = "SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM (SELECT a.LabelSeqNo, a.LabelFile, a.LabelPrinter, a.[Content], a.CreatedOn, a.PrintedOn," _
                    & " a.StatusCode, a.ErrMsg, row_number() OVER (partition BY a.LabelPrinter ORDER BY a.CreatedOn) AS 'n'" _
                    & " FROM (SELECT LabelSeqNo, LabelFile, LabelPrinter, [Content], CreatedOn, PrintedOn, StatusCode, ErrMsg" _
                    & " FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1' and LabelFile not in ('CLMatLabel.lab') " _
                    & " GROUP BY LabelSeqNo, LabelPrinter, CreatedOn, LabelFile, [Content]," _
                    & " CreatedOn, PrintedOn, StatusCode, ErrMsg) a) b WHERE n <= 10 ORDER BY CreatedOn"

        cmdReadHeader = New SqlClient.SqlCommand(getLabel, myConn)
        'cmdReadHeader = New SqlClient.SqlCommand("SELECT LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,PrintedOn,StatusCode FROM T_LabelPrint with (nolock)  WHERE StatusCode = '1'", myConn)


        Try
            myConn.Open()
            drHeader = cmdReadHeader.ExecuteReader()
            While drHeader.Read()
                mydatarow = Labels.NewRow()
                If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("LabelFile") = drHeader.GetValue(1)
                If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("LabelPrinter") = drHeader.GetValue(2)
                If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Content") = drHeader.GetValue(3)
                If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("CreatedOn") = drHeader.GetValue(4)
                If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("PrintedOn") = drHeader.GetValue(5)
                If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("StatusCode") = drHeader.GetValue(6)
                Labels.Rows.Add(mydatarow)
                If Not drHeader.GetValue(3) Is DBNull.Value Then
                    arryDetail = Split(mydatarow("Content").ToString, "^")
                    For i = 0 To UBound(arryDetail) Step 2
                        mydatarow = Contents.NewRow()
                        mydatarow("LabelSeqNo") = drHeader.GetValue(0)
                        mydatarow("title") = arryDetail(i)
                        mydatarow("detail") = arryDetail(i + 1)
                        Contents.Rows.Add(mydatarow)
                    Next
                End If
            End While
            drHeader.Close()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("PublicFunction-OpenPrintNoMatlLabel", "", ex.Message & ex.Source, "E")
            ex.Message().ToString()
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function
    ''' <summary>
    ''' Get DaughterBoardSN List by IntSN
    ''' </summary>
    ''' <param name="IntSN">internal serial No</param>
    ''' <returns>sn list</returns>
    ''' <remarks>shihong_lin 3/21/2019</remarks>
    Public Function GetDaughterBoardSN(ByVal IntSN As String) As String()
        Using da As DataAccess = GetDataAccess()
            Dim SnArray As String()
            Try
                Dim Sqlstr As String
                Dim ds As DataSet
                Sqlstr = String.Format("select   IntSN from T_WIPHeader  with(nolock) where MotherBoardSN =(select WIPID from T_WIPHeader with(nolock) where IntSN='{0}') and MotherBoardSN <>''", IntSN)
                ds = da.ExecuteDataSet(Sqlstr)
                If ds.Tables.Count > 0 Then
                    Dim dt As DataTable = ds.Tables(0)
                    ReDim SnArray(dt.Rows.Count - 1)
                    For i As Integer = 0 To dt.Rows.Count - 1
                        SnArray(i) = dt.Rows(i).Item("IntSN").ToString()
                    Next
                End If
            Catch ex As Exception
                ErrorLogging("PublicFunction-GetDautherBoardSN", "", "IntSN: " & IntSN & ex.Message & ex.StackTrace, "E")

            End Try
            Return SnArray
        End Using
    End Function
    'For PH SFC / OneToOne module only in eTraceUI login check  'Add by Yudy Liu      -- 03/18/2015
#Region "PHUserAuthentication"

    Public Function ValidateUser(ByVal Uname As String, ByVal UPwd As String, ByVal CheckPWD As Boolean) As String
        Try
            Dim userinfo As MembershipUser = Membership.GetUser(Uname)
            If userinfo Is Nothing Then
                Return "Username not found"
            Else
                If Not userinfo.IsApproved Then
                    Return "Account is not yet Approved"
                ElseIf userinfo.IsLockedOut Then
                    Return "Account has been lockedout"
                Else
                    If Not Membership.ValidateUser(Uname, UPwd) And CheckPWD = True Then
                        Return "Password is invalid"
                    End If
                End If
            End If
        Catch ex As Exception
            ErrorLogging("UserAuthentication-ValidateUser", Uname, ex.Message & ex.Source, "E")
            Return ex.Message
            'Return "Cannot establish connection to System User's Server."
        End Try
        Return "OK"
    End Function

#End Region

End Class

'' Add by Jackson Huang 11/26/2013
Public Enum ConnectionSetting
    eTraceDBConnString = 0
    eTraceOTOConnString = 1
End Enum

