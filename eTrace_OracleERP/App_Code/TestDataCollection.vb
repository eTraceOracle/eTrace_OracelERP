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
Imports System.Data.SqlClient
Imports System.Collections

Public Structure TestJK
    Dim TEST As String
End Structure

Public Structure WIPHeader
    Public IntSN As String
    Public Model As String
    Public PCBA As String
    Public DJ As String
    Public ProdLine As String
    Public CurrentProcess As String
    Public TestRound As Integer
    Public Result As String
    Public AllPassed As Boolean
    Public MotherBoardSN As String
End Structure

Public Structure StatusHeaderStructure
    Public Model As String
    Public PCBA As String
    Public PO As String
    Public Process As String
    Public Result As String
    Public IntSerial As String
    Public ExtSerial As String
    Public OperatorName As String
    Public CurrDate As Date
    Public Tester As String
    Public ProgramName As String
    Public ProgramVersion As String
    Public IPSNo As String
    Public IPSVersion As String
    Public SalesOrder As String
    Public CustPN As String
    Public CustRev As String
    'Public AddlData As String
    Public SN2 As String
    Public SN3 As String
    Public SN4 As String
    Public BoxID As String
    Public PalletID As String
End Structure

'Public Structure BoxInfo
'    Public CustomerPN As String
'    Public Customer As String
'    Public BoxQty As Integer
'    Public PalletID As String
'    Public PalletQty As Integer
'End Structure

Public Structure ShipInfo
    Public CustomerPN As String
    Public Customer As String
    Public BoxID As String
    Public BoxQty As Integer
    Public PalletID As String
    Public boxsize As String
    Public Model As String
    Public prodOrder As String
End Structure
Public Structure RuleDetail
    Public Model As String
    Public Desc As String
    Public BU As String
    Public Customer As String
    Public CPN As String
    Public SNpattern As String
    Public SN2pattern As String
    Public SN3pattern As String
    Public SN4pattern As String
    Public VoltageType As String
    Public Power As String
    Public Mainboard As String
    Public SpecialRequirement As String
    Public Boxsize As Integer
    Public Palletsize As Integer
    Public Revision As String
    Public Remarks As String
    Public ExtSNSameIntSN As Boolean
    Public ConfirmSN As Boolean
End Structure

Public Class TestDataCollection
    Inherits PublicFunction

    'WIP_JOB_STATUS 1 Unreleased 
    'WIP_JOB_STATUS 3 Released 
    'WIP_JOB_STATUS 4 Complete 
    'WIP_JOB_STATUS 5 Complete - No Charges 
    'WIP_JOB_STATUS 6 On Hold 
    'WIP_JOB_STATUS 7 Cancelled 
    'WIP_JOB_STATUS 8 Pending Bill Load 
    'WIP_JOB_STATUS 9 Failed Bill Load 
    'WIP_JOB_STATUS 10 Pending Routing Load 
    'WIP_JOB_STATUS 11 Failed Routing Load 
    'WIP_JOB_STATUS 12 Closed 
    'WIP_JOB_STATUS 13 Pending - Mass Loaded 
    'WIP_JOB_STATUS 14 Pending Close 
    'WIP_JOB_STATUS 15 Failed Close 
    'WIP_JOB_STATUS 16 Pending Scheduling 
    'WIP_JOB_STATUS 17 Draft 

#Region "CopyDJ"
    Public Function CopyDJInfo(ByVal orgcode As String, ByVal isdelOldData As Integer) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet()
            CopyDJInfo = False
            ds.Tables.Add("T_DJInfo")

            Dim Oda As OracleDataAdapter = da.Oda_Sele()
            Dim comm_submit As OracleCommand = da.OraCommand()

            Try

                Oda.SelectCommand.CommandType = CommandType.StoredProcedure
                Oda.SelectCommand.CommandText = "APPS.xxetr_wip_pkg.get_dj_info"

                Oda.SelectCommand.Parameters.Add("o_cur_data", OracleType.Cursor).Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = orgcode

                Oda.SelectCommand.Connection.Open()
                Oda.Fill(ds, "T_DJInfo")
                Oda.SelectCommand.Connection.Close()
            Catch oe As Exception
                Throw oe
                ErrorLogging("TDC-CopyDJ", "", oe.Message & oe.Source, "E")
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try




            Dim i As Integer
            Dim CLMasterSQLCommand As SqlClient.SqlCommand
            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
            myConn.Open()
            If (isdelOldData = 1) Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("delete from T_DJInfo", myConn)
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                CLMasterSQLCommand.ExecuteNonQuery()
            End If
            CLMasterSQLCommand = New SqlClient.SqlCommand("insert into  T_DJInfo (PO,OrgCode,Status,model,pcba,POQty) values (@DJ,@OrgCode,@Status,@Model,@PCBA,@Qty)", myConn)
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            CLMasterSQLCommand.Parameters.Add("@DJ", SqlDbType.VarChar, 50, "PO")
            CLMasterSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
            CLMasterSQLCommand.Parameters.Add("@Status", SqlDbType.VarChar, 50, "Status")
            CLMasterSQLCommand.Parameters.Add("@Model", SqlDbType.VarChar, 50, "Model")
            CLMasterSQLCommand.Parameters.Add("@PCBA", SqlDbType.VarChar, 50, "PCBA")
            CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Int, 10, "POQty")
            Try
                For i = 0 To ds.Tables("T_DJInfo").Rows.Count - 1
                    CLMasterSQLCommand.Parameters("@DJ").Value = ds.Tables("T_DJInfo").Rows(i)("PO")
                    CLMasterSQLCommand.Parameters("@OrgCode").Value = orgcode 'ds.Tables("T_DJInfo").Rows(i)("OrgCode")
                    CLMasterSQLCommand.Parameters("@Status").Value = ds.Tables("T_DJInfo").Rows(i)("Status")
                    CLMasterSQLCommand.Parameters("@Model").Value = ds.Tables("T_DJInfo").Rows(i)("Model")
                    CLMasterSQLCommand.Parameters("@PCBA").Value = ds.Tables("T_DJInfo").Rows(i)("PCBA")
                    CLMasterSQLCommand.Parameters("@Qty").Value = ds.Tables("T_DJInfo").Rows(i)("POQty")
                    CLMasterSQLCommand.CommandType = CommandType.Text
                    CLMasterSQLCommand.ExecuteNonQuery()
                Next
                CopyDJInfo = True
                myConn.Close()
            Catch ex As Exception
                ErrorLogging("TDC-CopyDJ", "", ex.Message & ex.Source, "E")
            Finally
                If myConn.State <> ConnectionState.Closed Then myConn.Close()
            End Try
        End Using
    End Function



    Public Function CopyDJinfoUseServerName() As Boolean
        Dim servername As String = ""
        servername = System.Net.Dns.GetHostName()
        Dim cmdReadHeader As SqlClient.SqlCommand
        Dim drValue As SqlClient.SqlDataReader
        Dim index As Integer = 0
        CopyDJinfoUseServerName = False
        If (servername <> "") Then
            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            Try
                myConn.Open()
                'Dim CLMasterSQLCommand As SqlClient.SqlCommand
                'CLMasterSQLCommand = New SqlClient.SqlCommand("delete from T_DJInfo", myConn)
                'CLMasterSQLCommand.ExecuteNonQuery()

                cmdReadHeader = New SqlClient.SqlCommand("select Orgcode from T_Server where servername ='" + servername + "'", myConn)
                cmdReadHeader.CommandTimeout = TimeOut_M5
                drValue = cmdReadHeader.ExecuteReader()

                While drValue.Read
                    index += 1
                    CopyDJinfoUseServerName = CopyDJInfo(drValue.GetValue(0), index)
                End While
                myConn.Close()
            Catch ex As Exception
            Finally
                If myConn.State <> ConnectionState.Closed Then myConn.Close()
            End Try
        End If
    End Function


    Public Function isValidDJ(ByVal dj_name As String, ByVal status As String, ByVal OracleERPLogin As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet()
            isValidDJ = False
            ds.Tables.Add("isValidDJ")

            Dim oda As OracleDataAdapter = da.Oda_Sele
            Dim comm_submit As New OracleCommand()

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.xxetr_wip_pkg.validate_dj"

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_success_flag", OracleType.VarChar, 10))
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_error_mssg", OracleType.VarChar, 2000))

                oda.SelectCommand.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                oda.SelectCommand.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OracleERPLogin.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 50)).Value = dj_name
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_status_type", OracleType.VarChar, 50)).Value = status
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "isValidDJ")
                oda.SelectCommand.Connection.Close()
                Dim Flag, Msg As String
                Flag = oda.SelectCommand.Parameters("o_success_flag").Value.ToString
                Msg = oda.SelectCommand.Parameters("o_error_mssg").Value.ToString

                ' isValidDJ = isValidDJFromeTrace(dj_name)
                If (Flag = "Y") Then
                    isValidDJ = True
                End If
            Catch oe As Exception
                'isValidDJ = isValidDJFromeTrace(dj_name, OracleERPLogin.OrgCode)
                isValidDJ = False
                ErrorLogging("TDC-isValidDJ", OracleERPLogin.User, oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Private Function isValidDJFromeTrace(ByVal dj_name As String, ByVal OrgCode As String) As Boolean
        isValidDJFromeTrace = False
        Dim result As Object
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            myConn.Open()
            CLMasterSQLCommand = New SqlClient.SqlCommand("select count(*) from T_DJInfo where PO='" + dj_name + "' and OrgCode='" + OrgCode + "'", myConn)
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            result = CLMasterSQLCommand.ExecuteScalar
            If result.ToString <> "0" Then
                isValidDJFromeTrace = True
            End If
            myConn.Close()

        Catch ex As Exception
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

#End Region

#Region "test TD_header"
    Public Function insertTDHeader(ByVal startIndex As Integer, ByVal endIndex As Integer) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim myConn1 As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try

            Dim CLMasterSQLCommand As SqlClient.SqlCommand
            Dim myCommand As SqlClient.SqlCommand
            Dim ra As Integer

            Dim NextTDID As String = ""

            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn1)
            myCommand.CommandTimeout = TimeOut_M5
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.Add("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.Add("@TypeID", "TDID")

            myConn.Open()
            myConn1.Open()
            Dim myTrans As SqlClient.SqlTransaction = myConn.BeginTransaction()
            Dim myTransCommand As SqlClient.SqlCommand = myConn.CreateCommand()
            myTransCommand.Transaction = myTrans

            ' ra = myCommand.ExecuteNonQuery()
            'NextTDID = myCommand.Parameters(0).Value
            Dim sqlstr As String
            Dim i As Integer

            For i = startIndex To endIndex
                ra = myCommand.ExecuteNonQuery()
                NextTDID = myCommand.Parameters(0).Value
                sqlstr = String.Format("Insert into T_TDHeader2 ( TDID,ExtSerialNo,IntSerialNo,ProcessName,SeqNo,PO,Model,PCBA,ProdDate,Result,OperatorName,TesterNo,ProgramName,ProgramRevision,IPSNo,IPSRevision,Remark)" _
                                                                   & "values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')", _
                                                     NextTDID, _
                                                      "extseriaNO" + i.ToString, _
                                                      "IntSerial" + i.ToString, _
                                                      "Matching1", _
                                                      1, _
                                                     "PO" + i.ToString, _
                                                      "7001394-Y001", _
                                                      "7001394-Y001", _
                                                      Today, _
                                                      "Pass", _
                                                     "33974", _
                                                      "33974", _
                                                      DBNull.Value, _
                                                      DBNull.Value, _
                                                      DBNull.Value, _
                                                      DBNull.Value, _
                                                      DBNull.Value)
                '  CLMasterSQLCommand = New SqlClient.SqlCommand(sqlstr, myConn)
                myTransCommand.CommandText = sqlstr
                myTransCommand.ExecuteNonQuery()
                sqlstr = String.Format("Insert into T_TDHeader1 ( TDID,ExtSerialNo,IntSerialNo,ProcessName,SeqNo,PO,Model,PCBA,ProdDate,Result,OperatorName,TesterNo,ProgramName,ProgramRevision,IPSNo,IPSRevision,Remark)" _
                                                                  & "values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')", _
                                                    NextTDID, _
                                                     "extseriaNO" + i.ToString, _
                                                     "IntSerial" + i.ToString, _
                                                     "Matching1", _
                                                     1, _
                                                    "PO" + i.ToString, _
                                                     "7001394-Y001", _
                                                     "7001394-Y001", _
                                                     Today, _
                                                     "Pass", _
                                                    "33974", _
                                                     "33974", _
                                                     DBNull.Value, _
                                                     DBNull.Value, _
                                                     DBNull.Value, _
                                                     DBNull.Value, _
                                                     DBNull.Value)
                'CLMasterSQLCommand = New SqlClient.SqlCommand(sqlstr, myConn)
                'CLMasterSQLCommand.ExecuteNonQuery()
                myTransCommand.CommandText = sqlstr
                myTransCommand.ExecuteNonQuery()

            Next
            myTrans.Commit()
        Catch oe As Exception
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
            If myConn1.State <> ConnectionState.Closed Then myConn1.Close()
        End Try
        'myConn.Close()
    End Function
#End Region

#Region "Mac Address"

    ''''//-----------------------------------------------------------------------

    ''''//  The Mac Address assignment program will invoke the function as below.

    ''''//  <copyright>Copyright (c) EMRSN. All rights reserved.</copyright>

    ''''//  <date>03-14-2011</date>

    ''''//  <author>Jackson Huang</author>

    ''''//  <comment>

    ''''//  Corresponding datatable: T_LabelPrint, T_SEQNO

    ''''//  Corresponding store procedures: sp_PrintMacAdress, sp_RePrintMacAdress

    ''''//  </comment>

    ''''//-----------------------------------------------------------------------

    Public Function PrintMacAddress(ByVal macType As String, ByVal addressTotal As Integer, _
            ByVal user As String, ByVal printerName As String) As String

        ''Dim flag As Boolean = True
        Dim serverMsg As String = String.Empty
        Dim dbConnection As SqlConnection = Nothing
        Dim printComm As SqlCommand = Nothing
        Try
            dbConnection = New SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
            printComm = dbConnection.CreateCommand()
            dbConnection.Open()
            printComm.Transaction = dbConnection.BeginTransaction(IsolationLevel.ReadCommitted)

            ''dbConnection.Open()
            If (ConnectionState.Closed) Then
                printComm.Connection.Open()
            End If


            printComm.CommandType = CommandType.StoredProcedure
            printComm.CommandText = "sp_PrintMacAdress"

            printComm.Parameters.Add(New SqlParameter("@Type", SqlDbType.VarChar, 50)).Value = macType
            printComm.Parameters.Add(New SqlParameter("@Total", SqlDbType.Int)).Value = addressTotal
            printComm.Parameters.Add(New SqlParameter("@user", SqlDbType.VarChar, 50)).Value = user
            printComm.Parameters.Add(New SqlParameter("@printer", SqlDbType.VarChar, 50)).Value = printerName
            serverMsg = printComm.ExecuteScalar().ToString()
            If (serverMsg = "3") Then
                printComm.Transaction.Commit()

            Else
                printComm.Transaction.Rollback()

            End If
        Catch ex As NullReferenceException
            printComm.Transaction.Rollback()
            ErrorLogging("Print Mac Address", "", ex.Message & ex.Source, "E")
        Catch ex As Exception
            printComm.Transaction.Rollback()
            ErrorLogging("Print Mac Address", "", ex.Message & ex.Source, "E")
        Finally
            printComm.Connection.Close()
            'printComm.Connection.Dispose()
            'printComm.Dispose()
            dbConnection.Close()
            'dbConnection.Dispose()
        End Try


        Return serverMsg

    End Function

    Public Function ReprintMacAddress(ByVal macType As String, ByVal startAddress As String, _
            ByVal endAddress As String, ByVal user As String, ByVal printerName As String) As String

        ''Dim flag As Boolean = True
        Dim serverMsg As String = String.Empty
        Dim dbConnection As SqlConnection = Nothing
        Dim printComm As SqlCommand = Nothing


        Try
            dbConnection = New SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
            printComm = dbConnection.CreateCommand()
            dbConnection.Open()
            printComm.Transaction = dbConnection.BeginTransaction(IsolationLevel.ReadCommitted)

            ''dbConnection.Open()
            If (ConnectionState.Closed) Then
                printComm.Connection.Open()
            End If


            printComm.CommandType = CommandType.StoredProcedure
            printComm.CommandText = "sp_RePrintMacAdress"

            printComm.Parameters.Add(New SqlParameter("@Type", SqlDbType.VarChar, 50)).Value = macType
            printComm.Parameters.Add(New SqlParameter("@startHec", SqlDbType.VarChar, 4)).Value = startAddress
            printComm.Parameters.Add(New SqlParameter("@endHec", SqlDbType.VarChar, 4)).Value = endAddress
            printComm.Parameters.Add(New SqlParameter("@user", SqlDbType.VarChar, 50)).Value = user
            printComm.Parameters.Add(New SqlParameter("@printer", SqlDbType.VarChar, 50)).Value = printerName

            serverMsg = printComm.ExecuteScalar().ToString()
            If (serverMsg = "3") Then
                printComm.Transaction.Commit()
            Else
                printComm.Transaction.Rollback()

            End If

        Catch ex As NullReferenceException
            printComm.Transaction.Rollback()
            ErrorLogging("Reprint Mac Address", "", ex.Message & ex.Source, "E")
        Catch ex As Exception
            printComm.Transaction.Rollback()
            ErrorLogging("Reprint Mac Address", "", ex.Message & ex.Source, "E")
        Finally
            printComm.Connection.Close()
            'printComm.Connection.Dispose()
            'printComm.Dispose()
            dbConnection.Close()
            'dbConnection.Dispose()
        End Try


        Return serverMsg

    End Function

    Public Function GetPreMacAddress(ByVal AddressType As String) As String


        Using da As DataAccess = GetDataAccess()
            Dim Sqlstr As String
            Sqlstr = String.Format("SELECT Prefix FROM T_SEQNo WHERE TYPE = '{0}'", AddressType)
            Return Convert.ToString(da.ExecuteScalar(Sqlstr))
        End Using

    End Function

    Public Function GetLengthMacAddress(ByVal Category As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim Sqlstr As String
            Sqlstr = String.Format("SELECT Length FROM T_SEQNo WHERE TYPE = '{0}'", Category)
            Return Convert.ToString(da.ExecuteScalar(Sqlstr))
        End Using
    End Function

#End Region

#Region "Traveller Printing"

    Public Function GetTraveller(Optional ByVal model As String = "") As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                If (String.IsNullOrEmpty(model)) Then
                    strSQL = String.Format("SELECT * FROM T_Traveller")
                Else
                    strSQL = String.Format("SELECT * FROM T_Traveller WHERE Model = '{0}'", model)
                End If

                dt = da.ExecuteDataSet(strSQL, "Traveller").Tables(0)
            Catch ex As Exception
                ErrorLogging("GetTraveller", "", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function

    Public Function GetAllTravellerData() As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                strSQL = String.Format("SELECT Model, MaterialNo, Revision, WO, ECO, TVA, SubAssy, SubRevision, Operator1, Operator2, FontSize FROM T_Traveller")
                dt = da.ExecuteDataSet(strSQL, "Traveller").Tables(0)
            Catch ex As Exception
                ErrorLogging("GetTraveller", "", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function


    Public Function GetAllModels() As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                strSQL = String.Format("SELECT Model, Operator1, Operator2 FROM T_Traveller")
                dt = da.ExecuteDataSet(strSQL, "Traveller").Tables(0)
            Catch ex As Exception
                ErrorLogging("GetTraveller", "", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function


    Public Function UpdateTravellerInfo(ByVal ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                strSQL = String.Format("exec sp_UpdateTraveller '{0}'", DStoXML(ds))
                msg = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("TDC UpdateTravellerInfo", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function



    Public Function DeleteModel(ByVal model As String)
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("DELETE FROM T_Traveller WHERE Model='{0}'", model)
                da.ExecuteNonQuery(strSQL)
            Catch ex As Exception
                ErrorLogging("GetTraveller", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetPrinterServer() As String
        Dim result As String = String.Empty
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("SELECT Value FROM T_Config WHERE (ConfigID = '{0}') AND (Module = '{1}')", "PrinterServer", "SHOPFLOORCONTROL")
                result = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("GetPrinterServer", "", ex.Message & ex.Source, "E")
            End Try
        End Using
        Return result
    End Function

#End Region

#Region "QC Data Collection"

    ''' <summary>
    ''' Get Daughter Board By DJ
    ''' </summary>
    ''' <param name="DJ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDaughterBdByDJ(ByVal DJ As String, ByVal Flag As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL() As String
            Dim strResult As DataSet = New DataSet
            Try
                strSQL = New String() {String.Format("exec sp_GetDaughterBdByDJ '{0}', '{1}' ", DJ, Flag)}
                strResult = da.ExecuteDataSet(strSQL, New String() {"T1", "T2"})
            Catch ex As Exception
                ErrorLogging("GetDaughterBdByDJ", "MIC", ex.ToString, "E")
            End Try
            Return strResult
        End Using
    End Function

    Public Function GetSMTQCByDJ(ByVal DJ As String) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                strSQL = String.Format("sp_GetSMTQCByDJ '{0}'", DJ)
                da.ExecuteDataSet(strSQL)
                dt = da.ExecuteDataSet(strSQL).Tables(0)
            Catch ex As Exception
                ErrorLogging("QCDataCollection GetSMTQCByDJ", "MIC", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function


    Public Function GetSMTQCModels(ByVal lineType As String) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                strSQL = String.Format("sp_GetSMTQCModelsByLine '{0}'", lineType)
                dt = da.ExecuteDataSet(strSQL, "ModelInfo").Tables(0)
            Catch ex As Exception
                ErrorLogging("QCDataCollection GetSMTQCModels", "MIC", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function

    Public Function GetAllSMTQCData(ByVal ParamArray paras() As String) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dt As DataTable
            Try
                strSQL = String.Format("sp_GetSMTQCData '{0}', '{1}','{2}','{3}','{4}', '{5}', '{6}'", _
                paras(0).ToString(), paras(1).ToString(), paras(2).ToString(), paras(3).ToString(), paras(4).ToString(), paras(5).ToString(), paras(6).ToString())
                da.ExecuteDataSet(strSQL)
                ''dt = da.ExecuteDataSet("sp_GetSMTQCData", "SMTQCData").Tables(0)
                dt = da.ExecuteDataSet(strSQL).Tables(0)
            Catch ex As Exception
                ErrorLogging("QCDataCollection GetSMTQCModels", "MIC", ex.Message & ex.Source, "E")
            End Try
            Return dt
        End Using
    End Function




    Public Function SaveSMTQCData(ByVal dtSMTQC As DataTable, ByVal ParamArray paras() As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty

            Try

                Dim ds As DataSet = New DataSet("SMTQC")
                ds.Tables.Add(dtSMTQC)
                ''ErrorLogging("SFC SaveSMTQCData", "", DStoXML(ds))

                ''Dim xs As New System.Xml.Serialization.XmlSerializer(paras.GetType())
                ''Dim ms As New IO.MemoryStream()
                ''xs.Serialize(ms, paras)
                ''Dim resultXML As String = UTF8Encoding.UTF8.GetString(ms.ToArray())

                ''ErrorLogging("SFC SaveSMTQCData", "", DStoXML(ds) + " " + paras(8).ToString())

                strSQL = String.Format("exec sp_UpdateSMTQC N'{0}', '{1}', '{2}','{3}','{4}','{5}',N'{6}', '{7}', '{8}', '{9}'", _
                DStoXML(ds), paras(0).ToString(), paras(1).ToString(), paras(2).ToString(), paras(3).ToString(), paras(4).ToString(), paras(5).ToString(), paras(8).ToString(), paras(9).ToString(), paras(6).ToString())

                '@PModel			 varchar(50),
                '@PProdDate          DATETIME,
                '@PShift             varchar(50),
                '@PFloor             varchar(50),
                '@PLine              varchar(10),
                '@PUsername          varchar(50)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))
                ''ErrorLogging("SFC SaveSMTQCData", "", msg)

            Catch ex As Exception
                ErrorLogging("QCDataCollection SaveSMTQCData", "MIC", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function

    Public Function DeleteSMTQCData(ByVal ParamArray para() As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format(String.Format("sp_DeleteSMTQCData '{0}', '{1}', '{2}','{3}','{4}','{5}','{6}' ", para(0).ToString(), para(1).ToString(), para(2).ToString(), para(3).ToString(), para(4).ToString(), para(8).ToString(), para(6).ToString()))
                da.ExecuteNonQuery(strSQL)
            Catch ex As Exception
                ErrorLogging("QCDataCollection DeleteSMTQCData", "MIC", ex.Message & ex.Source, "E")
                Return "^TDC-366@"
            End Try
            Return "^TDC-49@"
        End Using
    End Function



#End Region

#Region "Smart Card Validation"
    'Public Function SaveSmartCardHistory(ByVal ParamArray cardParams As String()) As DataTable
    '    Using da As DataAccess = GetDataAccess()
    '        Dim strSQL, msg As String
    '        msg = String.Empty
    '        Try
    '            strSQL = String.Format("exec sp_SaveSmartCardHistory N'{0}', '{1}', '{2}','{3}','{4}','{5}',N'{6}', '{7}'", _
    '            cardParams(0), cardParams(1), cardParams(2), cardParams(3), cardParams(4), cardParams(5), cardParams(6), cardParams(7))
    '            ''msg = Convert.ToString(da.ExecuteScalar(strSQL))
    '            Dim dt As DataTable = da.ExecuteDataSet(strSQL, "SmartCardInfo").Tables(0)
    '            Return dt
    '        Catch ex As Exception
    '            ErrorLogging("TDC-WS-SaveSmartCardHistory", "", ex.Message & ex.Source, "E")
    '            Return New DataTable()
    '        End Try
    '        Return New DataTable()
    '    End Using
    'End Function


    Public Function GetSmartCardHistory(ByVal IntSN As String) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                strSQL = String.Format("exec sp_GetSmartCardHistory N'{0}'", _
                IntSN)
                ''msg = Convert.ToString(da.ExecuteScalar(strSQL))
                Dim dt As DataTable = da.ExecuteDataSet(strSQL, "SmartCardInfo").Tables(0)
                Return dt
            Catch ex As Exception
                ErrorLogging("TDC-WS-GetSmartCardHistory", "", ex.Message & ex.Source, "E")
                Return New DataTable()
            End Try
            Return New DataTable()
        End Using
    End Function

#End Region

#Region "Shop Floor Control"

#Region "Product Master"

    Public Function LookUpProductInfo() As DataSet

        ''Using da As DataAccess = GetDataAccess()
        ''    Try
        ''        Dim SqlProduct As String = String.Empty
        ''        Dim SqlTran As String = String.Empty
        ''        Dim SqlLabel As String = String.Empty
        ''        Dim SqlProcess As String = String.Empty
        ''        Dim SqlCPN As String = String.Empty
        ''        Dim SqlRoutingHeader As String = String.Empty

        ''        SqlProduct = String.Format("SELECT Model, Description, VotageType, Power, Remarks, Status, BusinessUnit, TraceabilityLevel FROM T_Product")
        ''        SqlTran = String.Format("SELECT Transmission, Description, TransmissionType, Recipient, Frequency, Status, DataPattern, Remarks FROM T_ProductTrans")
        ''        SqlLabel = String.Format("SELECT LabelID, Description, Type, SubType, NoOfUnit, NoOfPanel, Model, Customer, CustomerPN, Process, Path, SourceTable, Status FROM T_Label")
        ''        SqlProcess = String.Format("SELECT Process, Description, ProcessType, Status, TFS, ProdLineType, Equipment, CompTrace, Remarks FROM T_ProductProcess")
        ''        SqlCPN = String.Format("SELECT Customer, CustomerPN, Model FROM T_ProductCPN")
        ''        SqlRoutingHeader = String.Format("SELECT RoutingID, RoutingType, RoutingDesc, Remarks, StatusCode FROM T_RoutingHeader")

        ''        Dim sql() As String = {SqlProduct, SqlTran, SqlLabel, SqlProcess, SqlCPN, SqlRoutingHeader}
        ''        Dim tables() As String = {"ProductInfo", "TranInfo", "LabelInfo", "ProcessInfo", "CPNInfo", "RoutingHeader"}

        ''        Return da.ExecuteDataSet(sql, tables)     'dataset里有一个table,tablename is "PDTOItems".      
        ''    Catch ex As Exception
        ''        da.Dispose()
        ''        ErrorLogging("SFC LookUpProductInfo", "", ex.Message & ex.Source, "E")
        ''        ''Throw ex
        ''    End Try
        ''End Using

        Try

            Using da As DataAccess = GetDataAccess()
                Dim tables() As String = {"ProductInfo", "TranInfo", "LabelInfo", "ProcessInfo", "CPNInfo", "RoutingHeader"}
                Dim strSQL() As String = {String.Format("exec sp_GetPMDataByModel '{0}', {1}", String.Empty, 1)}
                Return da.ExecuteDataSet(strSQL, tables)
            End Using

        Catch ex As Exception
            ErrorLogging("SFC-LookUpProductInfo", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
        Return Nothing

    End Function


    '', ByVal msl As String
    Public Function LookUpSFCInfoByModel(ByVal model As String) As DataSet


        ''Using da As DataAccess = GetDataAccess()
        ''    Try
        ''        Dim SqlCPN, SqlStruct, SqlRouting, SqlProperties, SqlMSLHeader, SqlMSLItem As String

        ''        SqlCPN = String.Format("SELECT Model, Customer, CustomerPN, SerialNoPattern, SN2Pattern, SN3Pattern, SN4Pattern, BoxSize, PalletSize, Revision, DataTransmission, ExtSNSameIntSN, IPAddress, MacAddress, SNRangeControl, Remarks FROM T_ProductCPN WHERE Model = '{0}'", model)
        ''        SqlStruct = String.Format("SELECT Model, PCBA, Description, MotherBoard, IsTLA, Traced, RoutingID,Remarks FROM T_ProductStructure WHERE Model = '{0}'", model)
        ''        SqlRouting = String.Format("SELECT Model, PCBA, SeqNo, Process, SamplingTest, FailedUnitReturnTo, MaxTestRound, MaxFailure, StandardTime, SamplingLogic, Label1, Label2, PanelMatching, PanelSize, RetestFlag FROM T_ProductRouting WHERE Model = '{0}'", model)
        ''        SqlProperties = String.Format("SELECT Model, PCBA, SeqNo, PropertyType, PropertyName, InputType, Qty FROM T_ProductProperties WHERE Model = '{0}'", model)
        ''        SqlMSLHeader = String.Format("SELECT * FROM T_SFMSLHeader WHERE MSLID = '{0}'", model)
        ''        SqlMSLItem = String.Format("SELECT * FROM SqlMSLItem WHERE MSLID = '{0}'", model)

        ''        Dim sql() As String = {SqlCPN, SqlStruct, SqlRouting, SqlProperties}
        ''        Dim tables() As String = {"CPNInfo", "StructureInfo", "RoutingInfo", "PropertiesInfo"}

        ''        Return da.ExecuteDataSet(sql, tables)            'dataset里有两个table,tablename 是 "PDTOItems"和"test"
        ''    Catch ex As Exception
        ''        ErrorLogging("SFC LookUpSFCInfoByModel", "", ex.Message & ex.Source, "E")
        ''        ''Throw ex
        ''    End Try
        ''End Using

        Try

            Using da As DataAccess = GetDataAccess()
                Dim tables() As String = {"CPNInfo", "StructureInfo", "RoutingInfo", "PropertiesInfo", "MSLInfo"}
                Dim strSQL() As String = {String.Format("exec sp_GetPMDataByModel '{0}'", model)}
                Return da.ExecuteDataSet(strSQL, tables)
            End Using

        Catch ex As Exception
            ErrorLogging("SFC-LookUpSFCInfoByModel", "", "model: " & model & ", " & ex.Message & ex.Source, "E")
            Return Nothing
        End Try
        Return Nothing
    End Function

    Public Function LookUpRoutingInfoBy(ByVal routingId As String) As DataSet

        Dim ds As DataSet

        Using da As DataAccess = GetDataAccess()

            Try

                Dim sql() As String = {String.Format("exec sp_GetRoutingItem '{0}'", routingId), _
                      String.Format("exec sp_GetRoutingProperty '{0}'", routingId)}

                Dim nameArrary() As String = {"RoutingItem", "RoutingProperty"}

                ds = da.ExecuteDataSet(sql, nameArrary)

                Return ds           ''dataset里有两个table,tablename 是 "PDTOItems"和"test"
            Catch ex As Exception
                ErrorLogging("SFC LookUpRoutingInfoBy", "", ex.Message & ex.Source, "E")
                ''Throw ex
            End Try
        End Using

    End Function

    Public Function LookUpLabelPara(ByVal labelID As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Try
                Dim SqlLabelPara As String = String.Format("SELECT LabelID, Description, Parameter, Required, Type, Value, Prefix, Suffix FROM T_LabelPara WHERE LabelID = '{0}'", labelID)

                Dim sql() As String = {SqlLabelPara}
                Dim tables() As String = {"LabelParaInfo"}

                Return da.ExecuteDataSet(sql, tables)            'dataset里有两个table,tablename 是 "PDTOItems"和"test"
            Catch ex As Exception
                ErrorLogging("SFC LookUpLabelPara", "", ex.Message & ex.Source, "E")
                ''Throw ex
            End Try
        End Using
    End Function

    Public Function LookUpTransPara(ByVal transId As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Try
                Dim SqlLabelPara As String = String.Format("SELECT * FROM T_ProductTransParas WHERE Transmission = '{0}'", transId)

                Dim sql() As String = {SqlLabelPara}
                Dim tables() As String = {"TransParaInfo"}

                Return da.ExecuteDataSet(sql, tables)            'dataset里有两个table,tablename 是 "PDTOItems"和"test"
            Catch ex As Exception
                ErrorLogging("SFC LookUpTransPara", "", ex.Message & ex.Source, "E")

            End Try
        End Using

    End Function

    Public Structure ProductModel

        Public model As String
        Public description As String
        Public businessUnit As String
        Public votageType As String
        Public power As String
        Public changeBy As String
        Public status As String
        Public remarks As String
        Public username As String
        Public TraceabilityLevel As String

    End Structure

    Public Structure LabelModel
        Public labelID As String
        Public description As String
        Public type As String
        Public subType As String
        Public noUnit As String
        Public noPanel As String
        Public model As String
        Public customer As String
        Public cpn As String
        Public process As String
        Public path As String
        Public sourceTable As String
        Public status As String
    End Structure

    Public Structure ProcessEntity

        Public process As String
        Public description As String
        Public processType As String
        Public status As String
        Public remarks As String

    End Structure

    Public Function UploadProductInfo(ByVal productModel As ProductModel) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                strSQL = String.Format("exec sp_UpdateProduct '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}', '{8}'",
                productModel.model, productModel.description, productModel.businessUnit, productModel.votageType,
                productModel.power, productModel.changeBy, productModel.status, productModel.remarks, productModel.TraceabilityLevel)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC UploadProductInfo", "", ex.Message & ex.Source, "E")
                Return msg
            Finally
                da.Dispose()
            End Try

            Return msg
        End Using

    End Function
    ''' <summary>
    ''' clone an existing model
    ''' </summary>
    ''' <param name="newModel">new Model</param>
    ''' <param name="originalModel">original Model</param>
    ''' <param name="userName"></param>
    ''' <returns></returns>
    Public Function CloneProductInfo(ByVal newModel As String, ByVal originalModel As String, ByVal userName As String) As DataSet
        Dim ds As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                strSQL = String.Format(" exec sp_CloneProduct '{0}','{1}','{2}'", originalModel, newModel, userName)
                ds = da.ExecuteDataSet(strSQL)

            Catch ex As Exception
                ds = New DataSet()
                Dim dt As New DataTable()
                dt.Columns.Add("success")
                dt.Columns.Add("errormessage")
                dt.Rows(0)("success") = 0
                dt.Rows(0)("errormessage") = ex.Message
                ds.Tables.Add(dt)
                ErrorLogging("SFC UploadProductInfo", "", ex.Message & ex.Source, "E")
                Return ds
            Finally
                da.Dispose()
            End Try

            Return ds
        End Using

    End Function
    Public Function UploadProcessInfo(ByVal processEntity As DataTable, ByVal userName As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim msg As String = String.Empty


            Try
                Dim ds As DataSet = New DataSet("ProductProcess")
                ds.Tables.Add(processEntity)
                ''ErrorLogging("SFC UploadProductRouting", userName, DStoXML(ds))
                Dim strSQL As String = String.Format("exec sp_UpdateProccess '{0}', N'{1}'", DStoXML(ds), userName)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC UploadProcessInfo", userName, ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg

        End Using

        Return True
    End Function

    Public Function UploadTransInfo(ByVal labelArray As String(), ByVal transEntity As DataTable) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            Dim rowCPN As DataRow = Nothing
            msg = String.Empty

            Try
                Dim dsLable As DataSet = New DataSet("Transmission")
                dsLable.Tables.Add(transEntity)
                strSQL = String.Format("exec sp_UpdateTransmission '{0}', '{1}', '{2}','{3}', '{4}', '{5}','{6}', N'{7}', '{8}'", _
                labelArray(0), labelArray(1), labelArray(2), labelArray(3), labelArray(4), labelArray(5), labelArray(6), labelArray(7), DStoXML(dsLable))
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC UploadTransInfo", "", ex.Message & ex.Source, "E")
                Return ex.Message
            End Try

            Return msg
        End Using


    End Function

    Public Function UploadProductCPN(ByVal productCPN As DataTable, ByVal model As String, ByVal changedBy As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            Dim rowCPN As DataRow = Nothing
            msg = String.Empty
            If productCPN.Rows.Count > 0 Then
                Try

                    Dim dsCPN As DataSet = New DataSet("CPN")
                    dsCPN.Tables.Add(productCPN)
                    strSQL = String.Format("exec sp_InsertProductCPN '{0}', '{1}', {2}", DStoXML(dsCPN), model, changedBy)
                    msg = Convert.ToString(da.ExecuteScalar(strSQL))

                Catch ex As Exception
                    ErrorLogging("SFC UploadProductCPN", changedBy, ex.Message & ex.Source, "E")
                    Return msg
                End Try

                If (Not String.IsNullOrEmpty(msg)) Then
                    Return msg
                End If
            End If
            Return msg
        End Using

    End Function

    Public Function UploadProductStructure(ByVal productStructure As DataTable, ByVal model As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            Dim rowCPN As DataRow
            msg = String.Empty

            Try

                Dim dsStructure As DataSet = New DataSet("Structure")
                dsStructure.Tables.Add(productStructure)
                ''ErrorLogging("TDC-Upload Product Structure", "", DStoXML(dsStructure))
                strSQL = String.Format("exec sp_InsertProductStructure '{0}', '{1}'", DStoXML(dsStructure), model)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))


            Catch ex As Exception
                ErrorLogging("SFC UploadProductStructure", "", ex.Message & ex.Source, "E")
                da.Dispose()
                Return msg
            End Try
            Return msg
        End Using

    End Function

    Public Function UploadProductRouting(ByVal productRouting As DataTable, ByVal model As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            Dim rowCPN As DataRow
            Dim MaxTestRound As Int32 = 0
            Dim MaxFailure As Int32 = 0

            msg = String.Empty

            Try

                Dim dsRouting As DataSet = New DataSet("Routing")
                dsRouting.Tables.Add(productRouting)
                strSQL = String.Format("exec sp_InsertProductRouting '{0}', '{1}'", DStoXML(dsRouting), model)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC UploadProductRouting", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function

    Public Function UploadProductProperties(ByVal productProperties As DataTable, ByVal model As String) As String

        Using da As DataAccess = GetDataAccess()

            Dim flag As String = String.Empty
            Dim Sqlstr As String = String.Empty
            Dim strSQL As String = String.Empty
            Dim msg As String = String.Empty
            Dim qty As Int32 = 1
            Dim sign As Integer = 0

            Dim dr As DataRow = Nothing
            Try
                Dim dsProperties As DataSet = New DataSet("Properties")
                dsProperties.Tables.Add(productProperties)
                strSQL = String.Format("exec sp_InsertProductProperties '{0}', '{1}'", DStoXML(dsProperties), model)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception

                ErrorLogging("SFC UploadProductProperties", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using

    End Function

    Public Function UploadLabelInfo(ByVal labelModel As LabelModel, ByVal dtLabelPara As DataTable) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                Dim unit, panel As Integer
                unit = 0
                panel = 0
                Integer.TryParse(labelModel.noUnit, unit)
                Integer.TryParse(labelModel.noPanel, panel)
                strSQL = String.Format("exec sp_UpdateLabel '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'", _
                FixNull(labelModel.labelID), FixNull(labelModel.description), FixNull(labelModel.type), FixNull(labelModel.subType), FixNull(unit), FixNull(panel), FixNull(labelModel.model), _
                FixNull(labelModel.customer), FixNull(labelModel.cpn), FixNull(labelModel.process), FixNull(labelModel.path), FixNull(labelModel.sourceTable), FixNull(labelModel.status))
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

                If (Not String.IsNullOrEmpty(labelModel.labelID)) Then

                    Dim dsLable As DataSet = New DataSet("Label")
                    dsLable.Tables.Add(dtLabelPara)
                    strSQL = String.Format("exec sp_InsertLabelPara '{0}', '{1}'", DStoXML(dsLable), labelModel.labelID)
                    msg = Convert.ToString(da.ExecuteScalar(strSQL))

                End If

            Catch ex As Exception
                ErrorLogging("SFC UploadLabelInfo", "", ex.Message & ex.Source, "E")
                Return msg
            Finally
                da.Dispose()
            End Try
            Return msg
        End Using

        Return True

    End Function

    Public Function IsAuthenticateUser(ByVal userId As String) As Boolean
        Try
            Dim ds As DataSet = GetAllUsers(userId)
            If (Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            ErrorLogging("IsAuthenticateUser", "", ex.Message & ex.Source, "E")
            Return False
        End Try

    End Function

    Private Function GetAllUsers(ByVal userId As String) As DataSet

        Dim strSQL, msg As String
        msg = String.Empty

        Using da As DataAccess = GetDataAccess()
            Try
                strSQL = String.Format("SELECT UserID, Name, Division, Dept, Location, UserType, Status, LastLogon FROM T_Users WHERE (UserID = '{0}' OR EmpID = '{0}') AND Status = '{1}'", userId, "Active")
                Return da.ExecuteDataSet(strSQL, "UsersInfo")
            Catch ex As Exception
                ErrorLogging("GetAllUsers", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
        Return Nothing
    End Function

    Public Function UpdateRoutingInfo(ByVal dsRouting As DataSet, ByVal ParamArray list As Object()) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String = String.Empty
            Dim msg As String = String.Empty

            Dim dr As DataRow = Nothing
            Try
                ''Dim dsProperties As DataSet = New DataSet("Properties")
                dsRouting.DataSetName = "Routing"

                ''dsProperties.Tables.Add(productProperties)
                ''ErrorLogging("SFC UpdateRoutingInfo", "", DStoXML(dsRouting))
                strSQL = String.Format("exec sp_UpdateRoutingHeader '{0}','{1}','{2}','{3}','{4}','{5}'", _
                list(0).ToString(), list(1).ToString(), list(2).ToString(), list(3).ToString(), list(4).ToString(), list(5).ToString())
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

                ''ErrorLogging("SFC UpdateRoutingInfo", "", strSQL)
                strSQL = String.Format("exec sp_UpdateRoutingItem '{0}', '{1}'", DStoXML(dsRouting), list(0).ToString())
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception

                ErrorLogging("SFC UpdateRoutingInfo", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function

    Public Function GetProductLineInfo(Optional ByVal prodLine As String = "", Optional ByVal lineType As String = "") As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty

            Try
                Dim tablesName As String = "ProdLineInfo"
                strSQL = String.Format("exec sp_121PMGetProdLine")
                If (Not String.IsNullOrEmpty(prodLine) And String.IsNullOrEmpty(lineType)) Then
                    strSQL = String.Format("exec sp_121PMGetProdLine '{0}'", prodLine)
                ElseIf (Not String.IsNullOrEmpty(prodLine) And Not String.IsNullOrEmpty(lineType)) Then
                    strSQL = String.Format("exec sp_121PMGetProdLine '{0}', '{1}'", prodLine, lineType)
                End If


                Return da.ExecuteDataSet(strSQL, tablesName)

            Catch ex As Exception
                ErrorLogging("SFC GetProductLineInfo", "", ex.Message & ex.Source, "E")
                ''Return New DataTable("ProdLine")
            End Try
            Return Nothing
        End Using
    End Function


    Public Function UpdateProductLine(ByVal dt As DataTable, ByVal orgCode As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty

            Try
                strSQL = String.Format("exec sp_121PMUpdateProdLine '{0}', '{1}'", DTtoXML(dt), orgCode)
                'ErrorLogging("SFC-UpdateProductLine", "", DTtoXML(dt))
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC-UpdateProductLine", "", ex.Message & ex.Source, "E")
                msg = "^TDC"
                Return msg
            End Try
            Return msg
        End Using
    End Function

    Public Function GetEquipmentInfo(Optional ByVal equipmentId As String = "") As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), msg, tables() As String
            msg = String.Empty
            Try

                If (Not String.IsNullOrEmpty(equipmentId)) Then
                    tables = New String() {"EquipmentInfo", "EquipmentAccyInfo", "ProdLineInfo"}
                    strSQL = New String() {String.Format("exec sp_121PMGetEqt '{0}'", equipmentId), String.Format("exec sp_121PMGetProdLine")}
                    Return da.ExecuteDataSet(strSQL, tables)
                Else
                    tables = New String() {"EquipmentInfo", "ProdLineInfo"}
                    strSQL = New String() {String.Format("exec sp_121PMGetEqt"), String.Format("exec sp_121PMGetProdLine")}
                    Return da.ExecuteDataSet(strSQL, tables)
                End If

            Catch ex As Exception
                ErrorLogging("SFC-GetEquipmentInfo", "", ex.Message & ex.Source, "E")
            End Try
            Return Nothing
        End Using

    End Function


    Public Function UpdateEquipmentInfo(ByVal dt As DataTable, ByVal ParamArray para() As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String = String.Empty
            Dim msg As String = String.Empty

            Dim dr As DataRow = Nothing
            Try
                Dim ds As DataSet = New DataSet()
                ''ds.DataSetName = "Routing"

                ds.Tables.Add(dt)
                ''ErrorLogging("SFC-UpdateEquipmentInfo", "", DStoXML(ds))
                strSQL = String.Format("exec sp_121PMUpdateEqt '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}'", _
                   para(0).ToString(), para(1).ToString(), para(2).ToString(), para(3).ToString(), para(4).ToString(), para(5).ToString(), _
                   para(7).ToString(), para(8).ToString(), para(9).ToString(), para(10).ToString(), para(11).ToString(), _
                   para(12).ToString(), para(13).ToString(), para(14).ToString(), DStoXML(ds))
                msg = Convert.ToString(da.ExecuteScalar(strSQL))
                ''ErrorLogging("SFC UpdateRoutingInfo", "", strSQL)
                ''strSQL = String.Format("exec sp_UpdateEquipmentInfo '{0}', '{1}'", DStoXML(dsRouting), list(0).ToString())
                ''msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception

                ErrorLogging("SFC-UpdateEquipmentInfo", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using

    End Function

    Public Function UpdatePMdata(ByVal productModel As ProductModel, ByVal ds As DataSet) As String()
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg, flag As String
            msg = String.Empty
            flag = 0
            Try
                strSQL = String.Format("exec sp_UpdateProduct '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}', '{8}'", _
                productModel.model, productModel.description, productModel.businessUnit, productModel.votageType, _
                productModel.power, productModel.changeBy, productModel.status, productModel.remarks, productModel.TraceabilityLevel)
                msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                If (Not String.IsNullOrEmpty(msg)) Then
                    flag = 1
                End If

                If (ds.Tables("CPNInfo") IsNot Nothing) Then
                    strSQL = String.Format("exec sp_InsertProductCPN '{0}', '{1}', {2}", DTtoXML(ds.Tables("CPNInfo")), productModel.model, productModel.changeBy)
                    msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                    If (Not String.IsNullOrEmpty(msg) And flag = 0) Then
                        flag = 2
                    End If
                End If


                If (ds.Tables("StructureInfo") IsNot Nothing) Then
                    strSQL = String.Format("exec sp_InsertProductStructure '{0}', '{1}'", DTtoXML(ds.Tables("StructureInfo")), productModel.model)
                    msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                    If (Not String.IsNullOrEmpty(msg) And flag = 0) Then
                        flag = 3
                    End If
                End If


                If (ds.Tables("RoutingInfo") IsNot Nothing) Then

                    'ErrorLogging("SFC-UploadProductInfo", "", DTtoXML(ds.Tables("RoutingInfo")))
                    strSQL = String.Format("exec sp_InsertProductRouting '{0}', '{1}'", DTtoXML(ds.Tables("RoutingInfo")), productModel.model)
                    msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                    If (Not String.IsNullOrEmpty(msg) And flag = 0) Then
                        flag = 4
                    End If
                End If


                If (ds.Tables("PropertiesInfo") IsNot Nothing) Then
                    strSQL = String.Format("exec sp_InsertProductProperties '{0}', '{1}'", DTtoXML(ds.Tables("PropertiesInfo")), productModel.model)
                    msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                    If (Not String.IsNullOrEmpty(msg) And flag = 0) Then
                        flag = 5
                    End If
                End If


                If (ds.Tables("MSLInfo") IsNot Nothing) Then

                    Dim dv As DataView = ds.Tables("MSLInfo").DefaultView

                    Dim dtDistinct As DataTable = dv.ToTable(True, "MSLID", "Model", "PCBA", "Process", "Revision")
                    dtDistinct.TableName = "MSLHeaderInfo"

                    strSQL = String.Format("exec sp_121PMUpdateMSL '{0}', N'{1}', '{2}', '{3}'", productModel.model, productModel.changeBy, DTtoXML(ds.Tables("MSLInfo")), DTtoXML(dtDistinct))
                    msg = IIf(String.IsNullOrEmpty(msg), Convert.ToString(da.ExecuteScalar(strSQL)), msg)
                    If (Not String.IsNullOrEmpty(msg) And flag = 0) Then
                        flag = 11
                    End If
                End If


            Catch ex As Exception
                ErrorLogging("SFC-UploadProductInfo", "", ex.Message & ex.Source, "E")
                Return New String() {"Save failed", ""}
            Finally
                da.Dispose()
            End Try

            If (String.IsNullOrEmpty(msg)) Then
                flag = 0
            End If
            Return New String() {msg, flag}

        End Using

    End Function

    Public Function LookUpMSLInfo(ByVal model As String) As DataSet

        Try

            Using da As DataAccess = GetDataAccess()
                ''Dim tables() As String = {"MSLInfo"}
                Dim strSQL As String = String.Format("exec sp_GetPMDataByModel '{0}', {1}", model, 2)
                Return da.ExecuteDataSet(strSQL, "MSLInfo")
            End Using

        Catch ex As Exception
            ErrorLogging("SFC-LookUpMSLInfo", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
        Return Nothing

    End Function

    Public Function GetHRTrainingInfo(ByVal ATEMachine As String, ByVal employeeID As String) As String

        Dim da As SqlClient.SqlDataAdapter
        Dim comm As SqlClient.SqlCommand
        Dim dt As DataTable = New DataTable("HRTrainingInfo")

        If (String.IsNullOrEmpty(ATEMachine)) Then
            ErrorLogging("SFC-GetHRTrainingInfo", employeeID, "Call funcion failed: ATEMachine is required", "I")
            Return "ATEMachine is required"
        End If

        If (String.IsNullOrEmpty(employeeID)) Then
            ErrorLogging("SFC-GetHRTrainingInfo", ATEMachine, "Call funcion failed: employeeID is required", "I")
            Return "EmployeeID is required"
        End If

        Try
            Dim query As String = String.Format("select i_course, i_binding from [V_TrainingDataeTrace] WITH(NOLOCK) WHERE i_empno = '{0}' ORDER BY i_binding desc", employeeID)

            Using conn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("HRTraningDBConnString"))
                da = New SqlClient.SqlDataAdapter(query, conn)
                conn.Open()
                comm = New SqlClient.SqlCommand(query, conn)
                da.Fill(dt)
                Return DTtoXML(dt)

            End Using

        Catch ex As Exception
            ErrorLogging("SFC-GetHRTrainingInfo", ATEMachine, employeeID & ex.Message & ex.Source, "E")
            Return "Get HRTraining ERROR"
        End Try


    End Function

#End Region

    Public Function ReadRuleFromDB() As DataSet
        Dim i As Integer
        Dim dsRule As DataSet = New DataSet("dsRule")
        Dim dtRule As DataTable = New DataTable("dtRule")
        Dim dtStructure As DataTable = New DataTable("dtStructure")
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("select * from T_ProductMaster"), myConn)
            da.Fill(dtRule)
            Dim da1 As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("select * from T_ProductStructure"), myConn)
            da1.Fill(dtStructure)
            dsRule.Tables.Add(dtRule)
            dsRule.Tables.Add(dtStructure)
        Catch ex As Exception
            ErrorLogging("TDC-ReadRuleFromDB", "", ex.Message & ex.Source, "E")
        End Try
        Return dsRule
    End Function

    Public Function postProdMaster(ByVal pordMaster As RuleDetail, ByVal ProdStructure As DataSet, ByVal SAPLoginData As ERPLogin) As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim i As Integer

        Try
            strSQL = String.Format("select count(*) from T_ProductMaster where Customer='{0}' and CustomerPN='{1}'", pordMaster.Customer, pordMaster.CPN)
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
            i = Convert.ToInt32(cmdSQL.ExecuteScalar)
            If i > 0 Then
                strSQL = String.Format("update T_ProductMaster set Model='{0}',Description='{1}',BusinessUnit='{2}'," _
                & "SerialNoPattern='{3}',SN2Pattern='{4}',SN3Pattern='{5}',SN4Pattern='{6}',VoltageType='{7}',Power='{8}',MainBoard='{9}'," _
                & "ChangedBy='{10}',ChangedOn=getDate(),SpecialRequirement='{11}',Remarks='{12}',Boxsize='{13}',Palletsize='{14}',Revision='{15}', " _
                & "ExtSNsameIntSN={16} ,ConfirmSN={17} where Customer='{18}' and CustomerPN='{19}'", _
                 pordMaster.Model, pordMaster.Desc, pordMaster.BU, pordMaster.SNpattern, pordMaster.SN2pattern, pordMaster.SN3pattern, _
                pordMaster.SN4pattern, pordMaster.VoltageType, pordMaster.Power, pordMaster.Mainboard, SAPLoginData.User, pordMaster.SpecialRequirement, pordMaster.Remarks, _
                pordMaster.Boxsize, pordMaster.Palletsize, pordMaster.Revision, Convert.ToInt32(pordMaster.ExtSNSameIntSN), Convert.ToInt32(pordMaster.ConfirmSN), pordMaster.Customer, pordMaster.CPN)
                cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                cmdSQL.ExecuteNonQuery()
                myConn.Close()
                postProdStructure(ProdStructure, pordMaster.Model)
                Return "update"
            Else

                strSQL = String.Format("insert into T_ProductMaster(Model,Description,BusinessUnit,Customer,CustomerPN," _
                & "SerialNoPattern,SN2Pattern,SN3Pattern,SN4Pattern,VoltageType,Power,MainBoard,ChangedBy,ChangedOn,SpecialRequirement,Boxsize,Palletsize,Revision,Remarks,ExtSNSameIntSN,ConfirmSN) " _
                & "values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}',getdate(),'{13}','{14}','{15}','{16}','{17}',{18},{19})", _
                pordMaster.Model, pordMaster.Desc, pordMaster.BU, pordMaster.Customer, pordMaster.CPN, pordMaster.SNpattern, pordMaster.SN2pattern, pordMaster.SN3pattern, pordMaster.SN4pattern, _
                pordMaster.VoltageType, pordMaster.Power, pordMaster.Mainboard, SAPLoginData.User, pordMaster.SpecialRequirement, pordMaster.Boxsize, pordMaster.Palletsize, pordMaster.Revision, pordMaster.Remarks, Convert.ToInt32(pordMaster.ExtSNSameIntSN), Convert.ToInt32(pordMaster.ConfirmSN))
                cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                cmdSQL.ExecuteNonQuery()
                myConn.Close()
                postProdStructure(ProdStructure, pordMaster.Model)
                Return "insert"
            End If

        Catch ex As Exception
            ErrorLogging("TDC-postProdMaster", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    Private Function postProdStructure(ByVal ProdStructure As DataSet, ByVal Model As String) As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim i As Integer

        Try
            Dim strModel, strPCBA, strDescription, strRemarks As String
            strModel = Model
            strSQL = String.Format("select count(*) from T_ProductStructure where Model='{0}'", strModel)
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
            i = Convert.ToInt32(cmdSQL.ExecuteScalar)
            If i > 0 Then
                strSQL = String.Format("delete from T_ProductStructure where Model='{0}'", strModel)
                cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                cmdSQL.ExecuteNonQuery()
                myConn.Close()
            End If

            If ProdStructure.Tables("dtStructure").Rows.Count > 0 Then
                If myConn.State = ConnectionState.Closed Then
                    myConn.Open()
                End If
                For i = 0 To ProdStructure.Tables("dtStructure").Rows.Count - 1
                    strModel = ProdStructure.Tables("dtStructure").Rows(i)("Model").ToString
                    strPCBA = ProdStructure.Tables("dtStructure").Rows(i)("PCBA").ToString
                    strDescription = ProdStructure.Tables("dtStructure").Rows(i)("Description").ToString
                    strRemarks = ProdStructure.Tables("dtStructure").Rows(i)("Remarks").ToString
                    strSQL = String.Format("insert into T_ProductStructure(Model,PCBA,Description,Remarks) " _
                    & "values('{0}','{1}','{2}','{3}')", strModel, strPCBA, strDescription, strRemarks)
                    cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                    cmdSQL.ExecuteNonQuery()
                Next
                myConn.Close()
            End If

        Catch ex As Exception
            ErrorLogging("TDC-postProdStructure", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    Public Function DelProdMaster(ByVal Model As String) As Boolean
        Dim dQty As Double
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim i As Integer
        Dim strSQL As String
        Try
            strSQL = String.Format("select count(*) from T_Shippment where Model='{0}'", Model)
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
            i = Convert.ToInt32(cmdSQL.ExecuteScalar)
            If i > 0 Then
                myConn.Close()
                Return False
            Else
                Dim strCMD As String
                strCMD = String.Format("Delete T_ProductMaster WHERE Model='{0}'", Model)
                cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
                cmdSQL.ExecuteNonQuery()
                myConn.Close()
                Return True
            End If
        Catch ex As Exception
            ErrorLogging("TDC-DelProdMaster", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try

    End Function

    Public Function ModelDefined(ByVal ModelNo As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, result As String
            Try
                strSQL = String.Format("exec sp_ModelDefined '{0}'", ModelNo)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
                Return result
            Catch ex As Exception
                ErrorLogging("SFC-ModelDefined", "", ex.Message & ex.Source, "E")
                Return "^TDC-1@ ^TDC-5@"
            End Try
        End Using
    End Function

    Public Function RevinBox(ByVal BoxID As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim result As String = ""
        Try
            strSQL = String.Format("exec sp_WIPPackingRevinBox '{0}'", BoxID)
            result = Convert.ToString(da.ExecuteScalar(strSQL))
            Return result
        Catch ex As Exception
            ErrorLogging("SFC-RevinBox", BoxID, ex.Message, "E")
            Return ""
        End Try
    End Function



    Public Function StructureReadByPCBA(ByVal Model As String, ByVal PCBA As String, ByVal mode As Integer) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sql As String = String.Format("exec SP_121StructureReadByPCBA '{0}','{1}',{2}", Model, PCBA, mode)
                Return da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("SFC-StructureReadByPCBA", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function ModelStructure(ByVal ModelNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_getModelStructure '{0}'", ModelNo)
                DS = da.ExecuteDataSet(strSQL, "ModelStructure")
            Catch ex As Exception
                ErrorLogging("SFC-ModelStructure", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function

    Public Function ProdQty(ByVal DJ As String, ByVal PCBA As String, ByVal OrgCode As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim iProdQty As Integer
            Try
                strSQL = String.Format("exec sp_getProdQty '{0}','{1}','{2}'", DJ, PCBA, OrgCode)
                iProdQty = Convert.ToInt32(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-ProdQty", "", ex.Message & ex.Source, "E")
            End Try
            Return iProdQty
        End Using
    End Function

    Public Function GetDJMatchedQty(ByVal DJ As String, ByVal PCBA As String, ByVal OrgCode As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim iProdQty As Integer
            Try
                strSQL = String.Format("exec sp_GetDJMatchedQty '{0}','{1}','{2}'", DJ, PCBA, OrgCode)
                iProdQty = Convert.ToInt32(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-GetDJMatchedQty", "", ex.Message & ex.Source, "E")
            End Try
            Return iProdQty
        End Using
    End Function

    Public Function InsertPoQty(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_InsertPoQty '{0}','{1}','{2}','{3}',{4}", OrgCode, DJ, Model, PCBA, Qty)
                da.ExecuteNonQuery(strSQL)
                Return True
            Catch ex As Exception
                ErrorLogging("SFC-InsertPoQty", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function CountPoQty(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer, ByVal ModelRev As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_InsertPoQty '{0}','{1}','{2}','{3}',{4},'{5}'", OrgCode, DJ, Model, PCBA, Qty, ModelRev)
                da.ExecuteNonQuery(strSQL)
                Return True
            Catch ex As Exception
                ErrorLogging("SFC-InsertPoQty", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function CountPoQtyII(ByVal OrgCode As String, ByVal DJ As String, ByVal Model As String, ByVal PCBA As String, ByVal Qty As Integer, ByVal ModelRev As String, ByVal DJType As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_CoutnPoQtyII '{0}','{1}','{2}','{3}',{4},'{5}','{6}'", OrgCode, DJ, Model, PCBA, Qty, ModelRev, DJType)
                da.ExecuteNonQuery(strSQL)
                Return True
            Catch ex As Exception
                ErrorLogging("SFC-InsertPoQty", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function IntSNIsValid(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_IntSNIsValid '{0}'", IntSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-IntSNIsValid", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function ReadPOQtyByPOAndPCBA(ByVal DJ As String, ByVal PCBA As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_ReadPOQtyByPOAndPCBA '{0}','{1}'", DJ, PCBA)
                DS = da.ExecuteDataSet(strSQL, "dsPOQty")
            Catch ex As Exception
                ErrorLogging("SFC-PCBARouting", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function


    Public Function PCBARouting(ByVal ModelNo As String, ByVal PCBA As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_getPCBARouting '{0}','{1}'", ModelNo, PCBA)
                DS = da.ExecuteDataSet(strSQL, "PCBARouting")
            Catch ex As Exception
                ErrorLogging("SFC-PCBARouting", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function

    Public Function PCBListOfRework(ByVal WIPID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_121PCBListOfRework '{0}'", WIPID)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-PCBListOfRework", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function WIPMatching1(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPMatching1 N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPMatching1", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    'Public Function WIPMatchingN(ByVal DSWIP As DataSet) As String
    '    Using da As DataAccess = GetDataAccess()
    '        Dim strSQL As String
    '        Dim result As String = ""
    '        Try
    '            strSQL = String.Format("exec sp_WIPMatchingN N'{0}'", DStoXML(DSWIP))
    '            result = da.ExecuteScalar(strSQL).ToString
    '        Catch ex As Exception
    '            ErrorLogging("SFC-WIPMatching1", "", ex.Message & ex.Source, "E")
    '        End Try
    '        Return result
    '    End Using
    'End Function
    Public Function WIPVisualInspection(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPVisualInspection N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPVisualInspection", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function WIPBurnIn(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPBurnIn N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPBurnIn", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function WIPIDSwop(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPIDSwop N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPIDSwop", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function WIPModelSwop(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPModelSwop N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPModelSwop", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function
    Public Function ComponentReplacement(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ComponentReplacement N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-ComponentReplacement", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function getPCBAinWIPHeader(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_getPCBAinWIPHeader '{0}'", IntSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-getPCBAinWIPHeader", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function MI_getPCBAinWIPHeader(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_121getPCBAinWIPHeader '{0}'", IntSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-MI_getPCBAinWIPHeader", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function DBoardIsValid(ByVal intSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_DBoardIsValid '{0}'", intSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-DBoradIsValid", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function ModelConfiguratorSNValid(ByVal SN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ModelConfigurator_SNValid '{0}'", SN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-DBoradIsValid", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function BackToEeprom(ByVal IntSN As String, ByVal ExtSN As String, ByVal Attribute As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_BackToEeprom '{0}','{1}','{2}'", IntSN, ExtSN, Attribute)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-BackToEeprom", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function MatchingAccount(ByVal DJ As String, ByVal OrgCode As String, ByVal PCBA As String, ByVal Allow As Integer, ByVal PanelSize As Integer) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                Dim strSQL As String
                Dim MatchSucessful As Integer
                strSQL = String.Format("exec sp_UpdateMatchingQty '{0}','{1}','{2}',{3},{4}", DJ, OrgCode, PCBA, Allow, PanelSize)
                MatchSucessful = Convert.ToInt32(da.ExecuteScalar(strSQL))
                If MatchSucessful = 0 Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                ErrorLogging("SFC-MatchingAccount", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function GetResult(ByVal SCID As String, ByVal Process As String) As String
        Using da As DataAccess = GetDataAccess()
            GetResult = ""
            Try
                Dim strSQL As String
                strSQL = String.Format("exec sp_GetResult '{0}','{1}'", SCID, Process)
                GetResult = da.ExecuteScalar(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetResult", "", ex.Message & ex.Source + "," + SCID + "," + Process, "E")
            End Try
        End Using
    End Function

    Public Function DJinBox(ByVal DJ As String, ByVal BoxID As String) As String
        Using da As DataAccess = GetDataAccess()
            DJinBox = ""
            Try
                Dim strSQL As String
                strSQL = String.Format("exec sp_WIPPackingDJinBox '{0}','{1}'", DJ, BoxID)
                DJinBox = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-DJinBox", "", ex.Message + "," + DJ + "," + BoxID)
            End Try
        End Using
    End Function

    Public Function CheckPrevResult(ByVal IntSerialNo As String, ByVal Process As String) As String
        Using da As DataAccess = GetDataAccess()
            'CheckPrevResult = ""
            Try
                Dim strSQL As String
                strSQL = String.Format("exec sp_CheckPrevResult '{0}','{1}'", IntSerialNo, Process)
                CheckPrevResult = da.ExecuteScalar(strSQL).ToString.Trim.ToUpper
            Catch ex As Exception
                CheckPrevResult = ex.Message
                ErrorLogging("SFC-CheckPrevResult", "", ex.Message & ex.Source + "," + IntSerialNo + "," + Process, "E")
            End Try
        End Using
    End Function

    Public Function GetProductCPNbyModel(ByVal Model As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetProductCPNbyModel '{0}'", Model)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetProductCPNbyModel", "", ex.Message & ex.Source)
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function LargeThanMaxTest(ByVal IntSerialNo As String, ByVal Process As String) As String
        Dim size As Integer = 0
        Using da As DataAccess = GetDataAccess()
            LargeThanMaxTest = "INVALID"
            Try
                Dim strSQL As String
                strSQL = String.Format("exec sp_LargeThanMaxTest '{0}','{1}'", IntSerialNo, Process)
                LargeThanMaxTest = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-LargeThanMaxTest", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    ''Public Function WIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal OperatorName As String) As Boolean
    ''    WIPIn = False
    ''    Using da As DataAccess = GetDataAccess()
    ''        Dim strSQL As String
    ''        Try
    ''            strSQL = String.Format("exec sp_WIPIn '{0}','{1}','{2}'", IntSerialNo, Process, OperatorName)
    ''            da.ExecuteNonQuery(strSQL)
    ''            WIPIn = True
    ''        Catch ex As Exception
    ''            ErrorLogging("SFC-WIPIn", OperatorName, ex.Message & ex.Source, "E")
    ''        End Try
    ''    End Using
    ''End Function

    Public Function WIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal OperatorName As String) As Boolean
        WIPIn = False
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPIn '{0}','{1}','{2}'", IntSerialNo, Process, OperatorName)
                If (da.ExecuteScalar(strSQL).ToString()) = "PASS" Then
                    WIPIn = True
                End If
            Catch ex As Exception
                ErrorLogging("SFC-WIPIn", OperatorName, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WIPOut(ByVal header As StatusHeaderStructure) As Boolean
        WIPOut = False
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPOut '{0}','{1}','{2}',{3}", header.IntSerial, header.Process, header.OperatorName, header.Result)
                da.ExecuteNonQuery(strSQL)
                WIPOut = True
            Catch ex As Exception
                ErrorLogging("SFC-WIPOut", header.OperatorName, ex.Message & ex.Source, "E")
                WIPOut = False
            End Try
        End Using
    End Function


    Public Function checkSamplingTest(ByVal IntSN, ByVal CurrProcess) As Boolean
        Using da As DataAccess = GetDataAccess()
            checkSamplingTest = False
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_CheckSamplingTest '{0}','{1}'", IntSN, CurrProcess)
                checkSamplingTest = Convert.ToBoolean(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-checkSamplingTest", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function ReadDBoards(ByVal header As StatusHeaderStructure) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_ReadDBoards '{0}','{1}'", header.IntSerial, header.Process)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-ReadDBoards", header.OperatorName, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function WIPOutMatchingN(ByVal header As StatusHeaderStructure, ByVal ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = "ERROR"
            Try
                strSQL = String.Format("exec sp_WIPOutMatchingN '{0}','{1}',N'{2}','{3}',N'{4}'", header.IntSerial, header.Process, header.OperatorName, header.Result, DStoXML(ds))
                result = da.ExecuteNonQuery(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPOutMatchingN", header.OperatorName, ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function


    Public Function GetDataByIntSN(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetWipHeaderRecord '{0}'", IntSN)
                GetDataByIntSN = da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("SFC-GetDataByIntSN", IntSN, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function RDCBoardSNValid(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCBoardSNValid '{0}'", IntSN)
                RDCBoardSNValid = da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("SFC-RDCBoardSNValid", IntSN, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function MatListOnPCBA(ByVal WIPID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_121MatListOnPCBAbyWIPID '{0}'", WIPID)
                Return da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("SFC-MatListOnPCBA", WIPID, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function IDSwop(ByVal header As StatusHeaderStructure, ByVal type As Integer) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_IDSwop '{0}','{1}','{2}',N'{3}',{4}", header.IntSerial, header.ExtSerial, header.Process, header.OperatorName, type)
                result = da.ExecuteNonQuery(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-IDSwop", header.OperatorName, header.IntSerial + ": " + ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function


    Public Function IDUpdate(ByVal IntSN As String, ByVal DJ As String, ByVal Model As String, ByVal TVANo As String, ByVal OrgCode As String, ByVal user As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_IDUpdate '{0}','{1}','{2}','{3}','{4}','{5}'", IntSN, DJ, Model, TVANo, OrgCode, user)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-IDUpdate", user, IntSN + ": " + ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function IntSNRecycle(ByVal header As StatusHeaderStructure) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_IntSNRecycle '{0}',N'{1}'", header.IntSerial, header.OperatorName)
                da.ExecuteNonQuery(strSQL)
                IntSNRecycle = True
            Catch ex As Exception
                ErrorLogging("SFC-IntSNRecycle", header.OperatorName, header.IntSerial + ": " + ex.Message & ex.Source, "E")
                IntSNRecycle = False
            End Try
            Return IntSNRecycle
        End Using
    End Function

    Public Function IntSNRecycleII(ByVal header As StatusHeaderStructure) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_IntSNRecycleII '{0}',N'{1}'", header.IntSerial, header.OperatorName)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-IntSNRecycleII", header.OperatorName, header.IntSerial + ": " + ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function GetProductCPN(ByVal CPN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetProductCPN '{0}'", CPN)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetProductCPN", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetResultList(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetResultList '{0}'", IntSN)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetResultList", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetNextProcess(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetResultList '{0}'", IntSN)
                Dim ds As DataSet = da.ExecuteDataSet(strSQL)
                ErrorLogging("SFC-GetNextProcess", "", DStoXML(ds), "E")
                If (ds IsNot Nothing) Then
                    Dim dt = da.ExecuteDataSet(strSQL).Tables(0)
                    Dim dr() = dt.Select(String.Format("Result='{0}'", "FAIL"))
                    If (dr.Length <> 0) Then
                        Return dr(0)("Process").ToString()
                    Else
                        dr = dt.Select(String.Format("Result='{0}'", ""))
                        Return dr(0)("Process").ToString()
                    End If
                    ''Return da.ExecuteDataSet(strSQL)
                End If

            Catch ex As Exception
                ErrorLogging("SFC-GetNextProcess", "", ex.Message & ex.Source, "E")
                Return "The smart card not existed in system, please double check it correct"
            End Try
        End Using
        Return "The smart card not existed in system, please double check it correct"
    End Function


    Public Function GetResultAndAttributesList(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetResultAndAttributesList '{0}'", IntSN)
                Return da.ExecuteDataSet(strSQL, 1)
            Catch ex As Exception
                ErrorLogging("SFC-GetResultAndAttributesList", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    'Public Function GetResultAndPCBAList(ByVal IntSN As String, ByVal Proc As String) As DataSet
    '    Using da As DataAccess = GetDataAccess()
    '        Dim strSQL As String
    '        Try
    '            strSQL = String.Format("exec sp_GetResultAndPCBAList '{0}','{1}'", IntSN, Proc)
    '            Return da.ExecuteDataSet(strSQL, 1)
    '        Catch ex As Exception
    '            ErrorLogging("SFC-GetResultAndAttributesList", "", ex.Message & ex.Source, "E")
    '            Return New DataSet
    '        End Try
    '    End Using
    'End Function

    Public Function GetBoxInfo(ByVal BoxID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetBoxInfo '{0}'", BoxID)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetBoxInfo", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetLabel1(ByVal Model As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetLabel1 '{0}'", Model)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-GetBoxInfo", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetPackingListLabel(ByVal Model As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetPackingListLabel '{0}'", Model)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-GetPackingListLabel", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetBoxQtyInPallet(ByVal PalletID As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetBoxQtyInPallet '{0}'", PalletID)
                Return Convert.ToInt32(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-GetBoxQtyInPallet", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function
    Public Function GetAllLocks() As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetAllLocks")
                GetAllLocks = da.ExecuteDataset(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("SFC-GeAllLocks", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function
    Public Function UpdateLockByID(lockid As String, symptom As String, te As String, pe As String, pqe As String, other As String, pbr As String, remarks As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = "ERROR"
            Try
                strSQL = String.Format("exec sp_updateLockByID '{0}',N'{1}',N'{2}',N'{3}',N'{4}',N'{5}',N'{6}',N'{7}'", lockid, symptom, te, pe, pqe, other, pbr, remarks)
                ErrorLogging("SFC-UpdateLockByID", lockid, strSQL, "D")
                da.ExecuteNonQuery(strSQL)
                result = "true"
            Catch ex As Exception
                ErrorLogging("SFC-UpdateLockByID", lockid, ex.Message & ex.Source, "E")
                result = ex.Message
            End Try
            Return result
        End Using
    End Function
    Public Function UnlockdByID(lockid As String, user As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = "ERROR"
            Try
                strSQL = String.Format("exec sp_unLockByID '{0}',N'{1}'", lockid, user)
                da.ExecuteNonQuery(strSQL)
                result = "true"
            Catch ex As Exception
                ErrorLogging("SFC-UpdateLockByID", lockid, ex.Message & ex.Source, "E")
                result = ex.Message
            End Try
            Return result
        End Using
    End Function
    ''Public Function ATEWIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal User As String) As String
    ''    ATEWIPIn = ""
    ''    Dim returnValue As String
    ''    If ((IntSerialNo <> "") AndAlso (IntSNIsValid(IntSerialNo) = "^TDC-6@")) Then
    ''        ATEWIPIn = CheckPrevResult(IntSerialNo, Process)
    ''        If (ATEWIPIn.ToUpper = "PASS") Then
    ''            returnValue = LargeThanMaxTest(IntSerialNo, Process)
    ''            If (returnValue.ToUpper = "PASS") Then
    ''                ATEWIPIn = "PASS"
    ''                WIPIn(IntSerialNo, Process, User)
    ''            Else
    ''                ATEWIPIn = returnValue
    ''            End If
    ''            'Else
    ''            '    If (ATEWIPIn = "^SFC-7@") Then
    ''            '        ATEWIPIn = "Current porcess not define in flow."
    ''            '    ElseIf (ATEWIPIn = "^SFC-2@") Then
    ''            '        ATEWIPIn = "Current process already Pass."
    ''            '    ElseIf (ATEWIPIn = "^SFC-2@") Then
    ''            '        ATEWIPIn = "Prev Process Result is not pass."
    ''            '    End If
    ''        End If
    ''    Else
    ''        ATEWIPIn = "^SFC-6@"
    ''    End If
    ''End Function

    Public Function ATEWIPIn(ByVal IntSerialNo As String, ByVal Process As String, ByVal User As String) As String
        ATEWIPIn = ""
        Dim returnValue As String
        If ((IntSerialNo <> "") AndAlso (IntSNIsValid(IntSerialNo) = "^TDC-6@")) Then
            ATEWIPIn = CheckPrevResult(IntSerialNo, Process)
            If (ATEWIPIn.ToUpper = "PASS") Then
                returnValue = LargeThanMaxTest(IntSerialNo, Process)
                If (returnValue.ToUpper = "PASS") Then
                    If (WIPIn(IntSerialNo, Process, User)) Then
                        ATEWIPIn = "PASS"
                    Else
                        ATEWIPIn = "WIPINFAIL"
                    End If
                Else
                    ATEWIPIn = returnValue
                End If
            End If
        Else
            ATEWIPIn = "^TDC-154@"
        End If
    End Function

    'Public Function ATEWIPout(ByVal dswip As DataSet) As String
    '    Return ATEWIPout(dswip.GetXml)
    'End Function

    Public Function ATEWIPout(ByVal XML As String) As String
        'ErrorLogging("SFC-ATEWIPout", "Test", XML)
        ATEWIPout = ""
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPOutATE N'{0}'", XML)
                ATEWIPout = da.ExecuteScalar(strSQL).ToString

            Catch ex As Exception
                ErrorLogging("SFC-ATEWIPout", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function ATEWIPOutDirect(ByVal XML As String) As String
        'ErrorLogging("SFC-ATEWIPout", "Test", XML)
        ATEWIPOutDirect = ""
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPOutDirect N'{0}'", XML)
                ATEWIPOutDirect = da.ExecuteScalar(strSQL).ToString

            Catch ex As Exception
                ErrorLogging("SFC-ATEWIPoutDirect", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WIPPacking(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPPacking N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPPacking", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    'Public Function PrintPaking(ByVal CartonID As String, ByVal printer As String) As String
    '    Using da As DataAccess = GetDataAccess()
    '    End Using
    'End Function


    Public Function IsWipIn(ByVal IntSerial As String, ByVal CurrProcess As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As Integer
            Try
                strSQL = String.Format("exec sp_IsWipIn '{0}','{1}'", IntSerial, CurrProcess)
                result = Convert.ToInt32(da.ExecuteScalar(strSQL))
                If (result = 1) Then
                    IsWipIn = True
                Else
                    IsWipIn = False
                End If
            Catch ex As Exception
                IsWipIn = False
                ErrorLogging("SFC-IsWipIn", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function


    Public Function BurnInWipOut(ByVal Header As StatusHeaderStructure, ByVal checked As Boolean) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_BurnInWipOut '{0}','{1}','{2}','{3}','{4}',{5},'{6}'", Header.IntSerial, Header.Process, Header.OperatorName, Header.Result, Header.ProgramName, Header.ProgramVersion, Header.Tester)
                da.ExecuteNonQuery(strSQL)
                BurnInWipOut = True
            Catch ex As Exception
                BurnInWipOut = False
                ErrorLogging("SFC-BurnInWipOut", Header.OperatorName, ex.Message & ex.Source, "E")
            End Try

        End Using
    End Function

    Public Function GetBurnInTime(ByVal IntSerial As String, ByVal CurrProcess As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As DateTime
            Try
                strSQL = String.Format("exec sp_GetWipInTime '{0}','{1}'", IntSerial, CurrProcess)
                result = Convert.ToDateTime(da.ExecuteScalar(strSQL))
                GetBurnInTime = DateDiff("n", result, Now)
            Catch ex As Exception
                ErrorLogging("SFC-IsWipIn", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetShipInfoBySN(ByVal SerialNo As String) As ShipInfo
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As DataTable
            Dim newShipInfo As ShipInfo = New ShipInfo
            Try
                strSQL = String.Format("exec sp_GetShipInfoBySN '{0}'", SerialNo)
                result = da.ExecuteDataTable(strSQL)
                If (result.Rows.Count <> 0) Then
                    newShipInfo.CustomerPN = result.Rows(0)("CustomerPN").ToString
                    newShipInfo.Customer = result.Rows(0)("Customer").ToString
                    newShipInfo.PalletID = result.Rows(0)("PalletID").ToString
                    newShipInfo.BoxID = result.Rows(0)("BoxID").ToString
                    newShipInfo.Model = result.Rows(0)("Model").ToString
                    newShipInfo.prodOrder = result.Rows(0)("prodOrder").ToString
                    Return newShipInfo
                End If
            Catch ex As Exception
                ErrorLogging("SFC-GetShipInfoBySN", "", ex.Message & ex.Source, "E")
                Return newShipInfo
            End Try
        End Using
    End Function


    Public Function GetShipInfoByBoxID(ByVal BoxID As String, ByVal user As String) As ShipInfo
        Dim result As DataSet
        Dim shippmentinfo As DataSet
        Dim newShipInfo As ShipInfo = New ShipInfo
        Try
            result = GetBoxInfo(BoxID)
            shippmentinfo = getShipmentByBoxid(BoxID, user)
            If (result.Tables.Count > 0 AndAlso result.Tables(0).Rows.Count <> 0) Then
                newShipInfo.CustomerPN = result.Tables(0).Rows(0)("CustomerPN")
                newShipInfo.Customer = result.Tables(0).Rows(0)("Customer")
                newShipInfo.PalletID = result.Tables(0).Rows(0)("PalletID")
                newShipInfo.BoxID = result.Tables(0).Rows(0)("cartonid")
                newShipInfo.BoxQty = result.Tables(0).Rows(0)("packedQTY")
                newShipInfo.boxsize = result.Tables(0).Rows(0)("boxsize")
            End If
            If (shippmentinfo.Tables.Count > 0 AndAlso shippmentinfo.Tables(0).Rows.Count <> 0) Then
                newShipInfo.Model = shippmentinfo.Tables(0).Rows(0)("Model")
                newShipInfo.prodOrder = shippmentinfo.Tables(0).Rows(0)("prodOrder")
            End If
            Return newShipInfo
        Catch ex As Exception
            ErrorLogging("SFC-GetShipInfoByBoxID", "", ex.Message & ex.Source, "E")
            Return newShipInfo
        End Try
    End Function

    Public Function GetShipInfoByPalletID(ByVal PalletID As String, ByVal user As String) As ShipInfo
        Using da As DataAccess = GetDataAccess()

            Dim strSQL As String
            Dim result As DataTable
            Dim newShipInfo As ShipInfo = New ShipInfo
            Try
                strSQL = String.Format("exec sp_GetShipInfoByPalletID '{0}'", PalletID)
                result = da.ExecuteDataTable(strSQL)
                If (result.Rows.Count <> 0) Then
                    newShipInfo.CustomerPN = result.Rows(0)("CustomerPN")
                    newShipInfo.Customer = result.Rows(0)("Customer")
                    newShipInfo.PalletID = result.Rows(0)("PalletID")
                End If
                Return newShipInfo
            Catch ex As Exception
                ErrorLogging("SFC-GetShipInfoByPalletID", user, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function ChangePallet(ByVal BoxID As String, ByVal PalletID As String, ByVal user As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim strSQL As String
            Dim change As String
            Try
                strSQL = String.Format("exec sp_ChangePalletV2 '{0}','{1}','{2}'", BoxID, PalletID, user)
                change = Convert.ToString(da.ExecuteScalar(strSQL))
                Return change
            Catch ex As Exception
                ErrorLogging("SFC-ChangePallet", "", ex.Message & ex.Source, "E")
                Return "FAIL"
            End Try
        End Using
    End Function



    Public Function ChangeBox(ByVal SerialNo As String, ByVal oldboxid As String, ByVal newBoxID As String, ByVal user As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String
            Try
                strSQL = String.Format("exec sp_ChangeBoxV2 '{0}','{1}','{2}','{3}'", SerialNo, oldboxid, newBoxID, user)
                result = Convert.ToInt32(da.ExecuteScalar(strSQL))
                If (result = 0) Then
                    Return "^TDC-145@"
                ElseIf (result = 1) Then
                    Return "^TDC-103@"
                ElseIf (result = 2) Then
                    Return "^TDC-105@"
                ElseIf (result = 3 Or result = 4) Then
                    Return "^TDC-479@"
                End If
            Catch ex As Exception
                ErrorLogging("SFC-ChangeBox", user, ex.Message & ex.Source, "E")
                Return "FAIL"
            End Try
        End Using
    End Function

    Public Function OQAWipIn(ByVal ExtNo As String, ByVal OperatorName As String, ByVal InvOrg As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPInOQA '{0}','{1}','{2}'", ExtNo, OperatorName, InvOrg)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPPacking", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function ExistsFunctionalTest(ByVal ExtNo As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ExistsFunctionalTest '{0}'", ExtNo)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-ExistsFunctionalTest", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function TLAFlow(ByVal Model As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_TLAFlow '{0}'", Model)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-ExistsFunctionalTest", "", ex.Message & ex.Source, "E")
            End Try
            Return New DataSet
        End Using
    End Function

    Public Function getShipmentByBoxid(ByVal boxid As String, ByVal user As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_getShipmentByBoxid '{0}'", boxid)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-getShipmentByBoxid", user, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function OQACosmetic(ByVal ExtSN As String, ByVal Model As String, ByVal RetestNo As String, ByVal Result As String, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPOQACosmetic '{0}','{1}',N'{2}','{3}','{4}','{5}','{6}',N'{7}'", ExtSN, Model, ERPLoginData.User, ERPLoginData.OrgCode, ERPLoginData.ProductionLine, RetestNo, Result, DStoXML(dsFlow))
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-OQACosmetic", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function SFCDBoardIsValid(ByVal intSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_SFCDBoardIsValid '{0}'", intSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-DBoradIsValid", "", ex.Message)
            End Try
            Return result
        End Using
    End Function


    Public Function PCBAQtyOnDJ(ByVal IntSN As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim Qty As Integer
            Try
                strSQL = String.Format("exec sp_getDJQtyByIntSN '{0}'", IntSN)
                Qty = Convert.ToInt32(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-PCBAQtyOnDJ", "", ex.Message)
            End Try
            Return Qty
        End Using
    End Function

    Public Function Get_CusTableLists(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select * from T_CusTables")
                ds = da.ExecuteDataSet(Sqlstr, "CusTables")
                Return ds

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-Get_CusTableLists", LoginData.User.ToUpper, ex.Message)
                Return Nothing
            End Try
        End Using
    End Function


    ''Online Label Print
    '' Update by Jackson on 12/2/2013
    '' Fix the last character is '^' issue.
    Public Function PrintPaking(ByVal boxid As String, ByVal labelid As String, ByVal printer As String) As Boolean

        If (String.IsNullOrEmpty(boxid) Or String.IsNullOrEmpty(labelid) Or String.IsNullOrEmpty(printer)) Then
            ErrorLogging("SFC_LabelPrintOnline_OpenPrintLabels", boxid, labelid, "I")
            Return False
        End If

        Using da As DataAccess = GetDataAccess()
            PrintPaking = True
            Dim myDataColumn As DataColumn
            Dim mydatarow As DataRow
            Dim head As String = ""
            Dim value As String = ""
            Dim labelfile As String = ""
            Dim OpenPrintLabel As New DataSet
            Dim strSQL As String

            Dim Labels As DataTable
            Labels = New Data.DataTable("Label")
            myDataColumn = New Data.DataColumn("LabelID", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Description", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Type", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Model", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Customer", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Process", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Path", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
            Labels.Columns.Add(myDataColumn)
            OpenPrintLabel.Tables.Add(Labels)

            Dim LabelPara As DataTable
            LabelPara = New Data.DataTable("LabelPara")
            myDataColumn = New Data.DataColumn("parameter", System.Type.GetType("System.String"))
            LabelPara.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("type", System.Type.GetType("System.String"))
            LabelPara.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("value", System.Type.GetType("System.String"))
            LabelPara.Columns.Add(myDataColumn)


            OpenPrintLabel.Tables.Add(LabelPara)

            Dim drHeader As SqlClient.SqlDataReader
            Dim drDetail As SqlClient.SqlDataReader
            Dim dsValue As DataSet = New DataSet("dsValue")
            Dim content As String = String.Empty
            Try
                strSQL = String.Format("SELECT LabelID,Description ,Type,Model ,Customer,Process,Path,status from T_Label WITH(NOLOCK) where status = 'ACTIVE' and labelId= '{0}'", labelid)
                drHeader = da.ExecuteReader(strSQL)
                If (drHeader.Read()) Then
                    mydatarow = Labels.NewRow()
                    If Not drHeader.GetValue(0) Is DBNull.Value Then mydatarow("LabelID") = drHeader.GetValue(0)
                    If Not drHeader.GetValue(1) Is DBNull.Value Then mydatarow("Description") = drHeader.GetValue(1)
                    If Not drHeader.GetValue(2) Is DBNull.Value Then mydatarow("Type") = drHeader.GetValue(2)
                    If Not drHeader.GetValue(3) Is DBNull.Value Then mydatarow("Model") = drHeader.GetValue(3)
                    If Not drHeader.GetValue(4) Is DBNull.Value Then mydatarow("Customer") = drHeader.GetValue(4)
                    If Not drHeader.GetValue(5) Is DBNull.Value Then mydatarow("Process") = drHeader.GetValue(5)
                    If Not drHeader.GetValue(6) Is DBNull.Value Then mydatarow("Path") = drHeader.GetValue(6)
                    If Not drHeader.GetValue(7) Is DBNull.Value Then mydatarow("status") = drHeader.GetValue(7)
                    Labels.Rows.Add(mydatarow)
                    drHeader.Close()
                    labelfile = mydatarow("Path")
                End If


                strSQL = String.Format("Select parameter,type,value,prefix,suffix from T_LabelPara WITH(NOLOCK) where labelID ='{0}' and (type='Unique'  or type='Fix' or type='Count')", labelid)


                dsValue = da.ExecuteDataSet(strSQL, "Unique")



                Dim displayQty As String = String.Empty
                Dim size1 As Integer = 0
                Dim size2 As Integer = 0
                Dim type As String = ""
                Dim column As String = ""
                Dim column2 As String = ""
                Dim labelsize As Integer = 0
                Dim parameter As String = String.Empty

                Dim dr() As DataRow
                dr = dsValue.Tables(0).Select("parameter='SIZE1'")
                If (dr.Length > 0) Then
                    size1 = Integer.Parse(dr(0)("value").ToString)
                End If

                dr = dsValue.Tables(0).Select("parameter='SIZE2'")
                If (dr.Length > 0) Then
                    size2 = Integer.Parse(dr(0)("value").ToString)
                End If
                dr = dsValue.Tables(0).Select("parameter='QTY'")
                If (dr.Length > 0) Then
                    displayQty = dr(0)("Type").ToString().Trim()
                End If

                head = getValue(dsValue, labelid, boxid)
                If (String.IsNullOrEmpty(head)) Then
                    ErrorLogging("SFC_LabelPrintOnline_OpenPrintLabels", boxid, "The boxid does not have data in T_Shippment", "I")
                    Return False
                End If


                strSQL = String.Format("Select parameter,type,value ,prefix,suffix from T_LabelPara WITH(NOLOCK) where labelID ='{0}' and type='SN' ORDER BY value, Parameter", labelid)
                drDetail = da.ExecuteReader(strSQL)


                While drDetail.Read()
                    mydatarow = LabelPara.NewRow()

                    If (drDetail.GetValue(2) <> column2) Then
                        column = column2
                    End If
                    column2 = drDetail.GetValue(2)

                    mydatarow("parameter") = drDetail.GetValue(0)
                    mydatarow("type") = drDetail.GetValue(1)
                    type = drDetail.GetValue(1)
                    LabelPara.Rows.Add(mydatarow)
                End While

                If (column = "") Then
                    column = column2
                    column2 = ""
                    labelsize = LabelPara.Rows.Count
                Else
                    labelsize = LabelPara.Rows.Count / 2
                End If

                Dim nextid As Integer
                Dim index As Integer = 1

                Dim sn2 As ArrayList = New ArrayList
                Dim sns As ArrayList = getSNs(column, column2, labelid, boxid, sn2)
                Dim left As Integer = sns.Count Mod LabelPara.Rows.Count
                content = LabelPara.Rows(0)("parameter").ToString + "^" + sns(0)
                If (column2 <> "") Then
                    content = content + "^" + LabelPara.Rows(labelsize)("parameter").ToString.ToString + "^" + sn2(nextid)
                End If

                For nextid = 1 To sns.Count - 1
                    LabelPara.Rows(index)("value") = sns(nextid)

                    If (content <> "") Then
                        content = content + "^" + LabelPara.Rows(index)("parameter").ToString + "^" + sns(nextid)
                    Else
                        content = LabelPara.Rows(index)("parameter").ToString + "^" + sns(nextid)
                    End If

                    If (column2 <> "") Then
                        content = content + "^" + LabelPara.Rows(labelsize + index)("parameter").ToString + "^" + sn2(nextid)
                    End If
                    index = index + 1

                    If ((nextid + 1) Mod LabelPara.Rows.Count = 0) Then
                        content = head + content
                        If (displayQty <> "") Then
                            If (String.Equals(displayQty, "Unique")) Then
                                content = content + "^QTY^" + LabelPara.Rows.Count.ToString()
                            Else
                                content = content + "^QTY^" + sns.Count.ToString
                            End If

                        End If

                        If (size1 <> 0) Then
                            content = content + "^QTY1^" + size1.ToString
                        End If
                        If (size2 <> 0) Then
                            content = content + "^QTY2^" + (size2 - size1).ToString
                        End If
                        PrintLabels(printer, labelfile, content)
                        content = ""
                        index = 0
                    End If
                Next

                '' Construct empty value string, just like the box size is 20, only having one product in it, then print the 
                '' label
                Dim a1 As Integer = 0
                Dim a2 As Integer = 0
                If (content <> "") Then
                    content = head + content

                    Dim i As Integer = 0

                    For i = index To labelsize - 1
                        content = content + "^" + LabelPara.Rows(i)("parameter").ToString + "^"
                        If (column2 <> "") Then
                            content = content + "^" + LabelPara.Rows(labelsize + i)("parameter").ToString.ToString + "^"
                        End If
                    Next
                    If (displayQty <> "") Then
                        ''content = content + "^QTY^" + sns.Count.ToString ''+ "^"
                        If (String.Equals(displayQty, "Unique")) Then
                            content = content + "^QTY^" + (sns.Count.ToString Mod LabelPara.Rows.Count).ToString()
                        Else
                            content = content + "^QTY^" + sns.Count.ToString
                        End If
                    End If
                    If (size1 <> 0) Then
                        a1 = sns.Count Mod size2
                        If (a1 = 0) Then
                            a1 = size2
                        End If
                        If (a1 > size1) Then
                            a2 = a1 - size1
                            a1 = size1
                        Else
                            a2 = 0
                        End If
                        content = content + "^QTY1^" + a1.ToString + "^QTY2^" + a2.ToString
                    End If


                    PrintLabels(printer, labelfile, content)
                    ' PrintLabels(printer, labelfile, head + content + "QTY^" + sns.Count.ToString + "^")
                End If
            Catch ex As Exception
                PrintPaking = False
                ErrorLogging("SFC_LabelPrintOnline_OpenPrintLabels", boxid, ex.Message & ex.Source, "E")
                ErrorLogging("SFC_LabelPrintOnline_OpenPrintLabels", boxid, content, "I")
            End Try

        End Using
    End Function


    Public Function ExtractSN3(ByVal prodcutSN As String) As String

        '' Remove the leading 'S'
        If (prodcutSN.StartsWith("S", StringComparison.OrdinalIgnoreCase)) Then
            Return prodcutSN.Remove(0, 1)
        Else
            '' SN format: A9????0???? & SBR8???????
            If (prodcutSN.StartsWith("A", StringComparison.OrdinalIgnoreCase) Or prodcutSN.Contains("SBR8")) Then
                Return prodcutSN
            Else
                '' SN format: [)>061PC112567.B1A1TA9????0????18VLENSN4LCN16D????????
                '' SN format: SA9????0????
                If (prodcutSN.Contains("A9")) Then
                    Return prodcutSN.Substring(prodcutSN.IndexOf("A9"), 11)
                End If
            End If
        End If
        Return prodcutSN
    End Function



    Public Function PrintSNLabel3(ByVal prodcutSN As String, ByVal labelId As String, ByVal printerName As String) As String
        Try

            Dim dt As DataTable = Nothing
            Using da As DataAccess = GetDataAccess()
                Dim strSQL As String = String.Format("exec sp_121PMGetLabelInfo '{0}'", labelId)
                dt = da.ExecuteDataTable(strSQL)
            End Using

            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                dt.TableName = "LabelInfo"

                Dim labelFile As String = Trim(dt.Rows(0)("Path").ToString())

                If (String.IsNullOrEmpty(labelFile)) Then
                    ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, labelId)
                    Return "^TDC-30@"
                End If

                Dim content As StringBuilder = New StringBuilder()
                For Each row As DataRow In dt.Rows

                    Dim result As Boolean = IIf(String.IsNullOrEmpty(FixNull(row("Required"))), Boolean.FalseString, Boolean.FalseString)
                    Boolean.TryParse(row("Required").ToString(), result)
                    If (result) Then
                        content.Append(row("Parameter").ToString() + "^" + row("Value").ToString() + "^")
                    End If

                Next

                Dim ds As DataSet = New DataSet(), CPN As String = String.Empty
                ds.Tables.Add(dt)

                Dim tempDt As DataTable = ds.Tables(0).Clone(), flag As Boolean = False
                For Each row As DataRow In ds.Tables(0).Rows
                    If (Not String.Equals(row("Type").ToString(), "SN")) Then
                        tempDt.ImportRow(row)
                    ElseIf (String.Equals(row("Value").ToString(), "ProductSerialNO", StringComparison.OrdinalIgnoreCase)) Then

                        '' '' Remove the leading 'S'
                        ''If (prodcutSN.StartsWith("S", StringComparison.OrdinalIgnoreCase)) Then
                        ''    CPN = row("Parameter").ToString() + "^" + prodcutSN.Remove(0, 1) + "^"
                        ''Else
                        ''    '' SN format: A9????0???? & SBR8???????
                        ''    If (prodcutSN.StartsWith("A", StringComparison.OrdinalIgnoreCase)) Then
                        ''        CPN = row("Parameter").ToString() + "^" + prodcutSN + "^"
                        ''    Else
                        ''        '' SN format: [)>061PC112567.B1A1TA9????0????18VLENSN4LCN16D????????
                        ''        '' SN format: [)>0618VLELME1PBMK90591/121P???22PPSU,DC/DC,2715W12D????????SBR8???????
                        ''        If (prodcutSN.Contains("A9")) Then
                        ''            CPN = row("Parameter").ToString() + "^" + prodcutSN.Substring(prodcutSN.IndexOf("A9"), 11) + "^"
                        ''        ElseIf (prodcutSN.Contains("SBR8")) Then
                        ''            CPN = row("Parameter").ToString() + "^" + prodcutSN + "^"
                        ''        End If

                        ''    End If
                        ''End If

                        CPN = row("Parameter").ToString() + "^" + ExtractSN(prodcutSN) + "^"
                        flag = True
                    Else
                        Dim snArray As ArrayList = getSNs(row("Value").ToString(), String.Empty, labelId, String.Empty, New ArrayList(), prodcutSN)

                        '' '' Remove the S character.
                        ''If (snArray(0).StartsWith("S", StringComparison.OrdinalIgnoreCase)) Then
                        ''    snArray(0) = snArray(0).ToString().Remove(0, 1)
                        ''End If
                        If (Not String.IsNullOrEmpty(snArray(0))) Then
                            CPN = row("Parameter").ToString() + "^" + ExtractSN(snArray(0)) + "^"
                        End If

                    End If

                Next



                Dim tempDs As DataSet = New DataSet()
                tempDs.Tables.Add(tempDt)

                Dim IdxRevEnd As Integer = 0
                Dim IdxRevStart As Integer = -1

                '' Remove the Rev in CustomerPN
                Dim snStr As String = getValue(tempDs, labelId, String.Empty, prodcutSN)
                ''CPN += snStr
                IdxRevStart = snStr.IndexOf(".")
                If (IdxRevStart <> -1) Then
                    For i As Integer = snStr.IndexOf(".") To snStr.Length
                        If (String.Equals(snStr(i), "^")) Then
                            CPN = CPN + snStr.Remove(IdxRevStart, i - IdxRevStart)
                            Exit For
                        End If
                    Next
                Else
                    CPN += snStr
                End If

                '' Remove the last '^' character.
                CPN = CPN.Remove(CPN.Length - 1, 1)

                If (String.Equals(PrintLabels(printerName, labelFile, CPN), "True", StringComparison.OrdinalIgnoreCase)) Then
                    Return "^TDC-29@"
                Else
                    Return "^TDC-30@"
                End If

            Else
                ErrorLogging("SFC-WS-PrintSNLabel", labelId, "The label template can't find")
                Return "^TDC-30@"
            End If



        Catch ex As Exception
            ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, ex.Message & ex.Source)
            Return "^TDC-30@"
        End Try
    End Function


    '' To support vs2003
    Public Function ExtractSN(ByVal prodcutSN As String) As String

        prodcutSN = prodcutSN.ToUpper()
        '' Remove the leading 'S'
        If (prodcutSN.StartsWith("S")) Then
            Return prodcutSN.Remove(0, 1)
        Else
            '' SN format: A9????0???? & SBR8???????
            If (prodcutSN.StartsWith("A") Or prodcutSN.IndexOf("SBR8") <> -1) Then
                Return prodcutSN
            Else
                '' SN format: [)>061PC112567.B1A1TA9????0????18VLENSN4LCN16D????????
                '' SN format: SA9????0????
                If (prodcutSN.IndexOf("A9") <> -1) Then
                    Return prodcutSN.Substring(prodcutSN.IndexOf("A9"), 11)
                End If
            End If
        End If
        Return prodcutSN
    End Function

    '' To support vs2003
    Public Function PrintSNLabel4(ByVal prodcutSN As String, ByVal labelId As String, ByVal printerName As String) As String
        Try

            Dim dt As DataTable = Nothing
            Dim da As DataAccess = GetDataAccess()
            Dim strSQL As String = String.Format("exec sp_121PMGetLabelInfo '{0}'", labelId)
            dt = da.ExecuteDataTable(strSQL)

            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                dt.TableName = "LabelInfo"

                Dim labelFile As String = Trim(dt.Rows(0)("Path").ToString())

                If (labelFile = "") Then
                    ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, labelId)
                    Return "^TDC-30@"
                End If

                Dim content As String = String.Empty
                For Each row As DataRow In dt.Rows

                    ''Dim result As Boolean = IIf(String.IsNullOrEmpty(FixNull(row("Required"))), Boolean.FalseString, Boolean.FalseString)
                    Dim result As Boolean = IIf(FixNull(row("Required")) = "", Boolean.FalseString, Boolean.FalseString)
                    ''Boolean.TryParse(row("Required").ToString(), Result)
                    result = CType(row("Required").ToString(), Boolean)
                    If (result) Then
                        ''content.Append(row("Parameter").ToString() + "^" + row("Value").ToString() + "^")
                        content += row("Parameter").ToString() + "^" + row("Value").ToString() + "^"
                    End If

                Next

                Dim ds As DataSet = New DataSet(), CPN As String = String.Empty
                ds.Tables.Add(dt)

                Dim tempDt As DataTable = ds.Tables(0).Clone(), flag As Boolean = False
                For Each row As DataRow In ds.Tables(0).Rows
                    If (Not String.Equals(row("Type").ToString(), "SN")) Then
                        tempDt.ImportRow(row)
                    ElseIf (String.Equals(row("Value").ToString(), "ProductSerialNO")) Then

                        CPN = row("Parameter").ToString() + "^" + ExtractSN(prodcutSN) + "^"
                        flag = True
                    Else
                        Dim snArray As ArrayList = getSNs(row("Value").ToString(), String.Empty, labelId, String.Empty, New ArrayList(), prodcutSN)
                        If (Not String.IsNullOrEmpty(snArray(0))) Then
                            CPN = row("Parameter").ToString() + "^" + ExtractSN(snArray(0)) + "^"
                        End If

                    End If

                Next



                Dim tempDs As DataSet = New DataSet()
                tempDs.Tables.Add(tempDt)

                Dim IdxRevEnd As Integer = 0
                Dim IdxRevStart As Integer = -1

                '' Remove the Rev in CustomerPN
                Dim snStr As String = getValue(tempDs, labelId, String.Empty, prodcutSN)
                ''CPN += snStr
                IdxRevStart = snStr.IndexOf(".")
                If (IdxRevStart <> -1) Then
                    ''For i As Integer = snStr.IndexOf(".") To snStr.Length
                    ''    If (String.Equals(snStr(i), "^")) Then
                    ''        CPN = CPN + snStr.Remove(IdxRevStart, i - IdxRevStart)
                    ''        Exit For
                    ''    End If
                    ''Next
                    CPN = CPN + snStr.Remove(IdxRevStart, snStr.IndexOf("^", IdxRevStart) - IdxRevStart) ''snStr.IndexOf(".")
                Else
                    CPN += snStr
                End If

                '' Remove the last '^' character.
                CPN = CPN.Remove(CPN.Length - 1, 1)
                Dim AA As String = PrintLabels(printerName, labelFile, CPN)
                ErrorLogging("SFC-WS-PrintSNLabel", labelId, AA)
                If (String.Equals(PrintLabels(printerName, labelFile, CPN), "True", StringComparison.OrdinalIgnoreCase)) Then
                    Return "^TDC-29@"
                Else
                    Return "^TDC-30@"
                End If

            Else
                ErrorLogging("SFC-WS-PrintSNLabel", labelId, "The label template can't find")
                Return "^TDC-30@"
            End If



        Catch ex As Exception
            ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, ex.Message & ex.Source)
            Return "^TDC-30@"
        End Try
    End Function

    Public Function PrintSNLabel(ByVal prodcutSN As String, ByVal labelId As String, ByVal printerName As String) As String
        Try

            Dim dt As DataTable = Nothing
            Dim da As DataAccess = GetDataAccess()
            Dim strSQL As String = String.Format("exec sp_121PMGetLabelInfo '{0}'", labelId)
            dt = da.ExecuteDataTable(strSQL)

            If (Not dt Is Nothing AndAlso dt.Rows.Count > 0) Then
                dt.TableName = "LabelInfo"

                Dim labelFile As String = Trim(dt.Rows(0)("Path").ToString())

                If (labelFile = "") Then
                    ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, labelId)
                    Return "^TDC-30@"
                End If

                Dim content As String = String.Empty
                For Each row As DataRow In dt.Rows

                    ''Dim result As Boolean = IIf(String.IsNullOrEmpty(FixNull(row("Required"))), Boolean.FalseString, Boolean.FalseString)
                    Dim result As Boolean = IIf(FixNull(row("Required")) = "", Boolean.FalseString, Boolean.FalseString)
                    ''Boolean.TryParse(row("Required").ToString(), Result)
                    result = CType(row("Required").ToString(), Boolean)
                    If (result) Then
                        ''content.Append(row("Parameter").ToString() + "^" + row("Value").ToString() + "^")
                        content += row("Parameter").ToString() + "^" + row("Value").ToString() + "^"
                    End If

                Next

                Dim ds As DataSet = New DataSet, CPN As String = String.Empty
                ds.Tables.Add(dt)

                Dim tempDt As DataTable = ds.Tables(0).Clone(), flag As Boolean = False
                For Each row As DataRow In ds.Tables(0).Rows
                    If (Not String.Equals(row("Type").ToString(), "SN")) Then
                        tempDt.ImportRow(row)
                    ElseIf (String.Equals(row("Value").ToString(), "ProductSerialNO")) Then

                        CPN = row("Parameter").ToString() + "^" + ExtractSN(prodcutSN) + "^"
                        flag = True
                    Else
                        Dim snArray As ArrayList = getSNs(row("Value").ToString(), String.Empty, labelId, String.Empty, New ArrayList, prodcutSN)
                        ''If (Not String.IsNullOrEmpty(snArray(0))) Then
                        If (snArray(0) <> "") Then
                            CPN = row("Parameter").ToString() + "^" + ExtractSN(snArray(0)) + "^"
                        End If

                    End If

                Next



                Dim tempDs As DataSet = New DataSet
                tempDs.Tables.Add(tempDt)

                Dim IdxRevEnd As Integer = 0
                Dim IdxRevStart As Integer = -1

                '' Remove the Rev in CustomerPN
                Dim snStr As String = getValue(tempDs, labelId, String.Empty, prodcutSN)
                ''CPN += snStr
                IdxRevStart = snStr.IndexOf(".")
                If (IdxRevStart <> -1) Then
                    ''For i As Integer = snStr.IndexOf(".") To snStr.Length
                    ''    If (String.Equals(snStr(i), "^")) Then
                    ''        CPN = CPN + snStr.Remove(IdxRevStart, i - IdxRevStart)
                    ''        Exit For
                    ''    End If
                    ''Next
                    CPN = CPN + snStr.Remove(IdxRevStart, snStr.IndexOf("^", IdxRevStart) - IdxRevStart) ''snStr.IndexOf(".")
                Else
                    CPN += snStr
                End If

                '' Remove the last '^' character.
                CPN = CPN.Remove(CPN.Length - 1, 1)

                If (String.Equals(PrintLabels(printerName, labelFile, CPN), "True")) Then
                    Return "^TDC-29@"
                Else
                    Return "^TDC-30@"
                End If

            Else
                ErrorLogging("SFC-WS-PrintSNLabel", labelId, "The label template can't find")
                Return "^TDC-30@"
            End If



        Catch ex As Exception
            ErrorLogging("SFC-WS-PrintSNLabel", prodcutSN, ex.Message & ex.Source)
            Return "^TDC-30@"
        End Try
    End Function


    Private Function getValue5(ByVal dsvalue As DataSet, ByVal labelid As String, ByVal Boxid As String, Optional ByVal productSN As String = "") As String
        Using da As DataAccess = GetDataAccess()
            getValue5 = ""
            Dim drValue As DataSet = New DataSet
            Dim i As Integer
            Dim values As String = ""
            Dim value As String = ""
            Dim dr As DataRow() = dsvalue.Tables(0).Select("Type='Unique'")
            Dim strSQL As String = ""

            If (dr.Length = 0) Then
                Return ""
            End If

            If (dr.Length = 1) Then
                values = dr(0)("Value")
            Else
                For i = 0 To dr.Length - 2
                    values = values + dr(i)("Value") + ","
                Next
                values = values + dr(dr.Length - 1)("Value")
            End If


            Try

                If (String.IsNullOrEmpty(productSN)) Then
                    strSQL = String.Format("Select distinct {0} from V_PackingLabelPrint  where  cartonid='{1}' ", values, Boxid)
                Else
                    strSQL = String.Format("Select distinct {0} from V_PackingLabelPrint  where  ProductSerialNo='{1}' ", values, productSN)
                End If

                drValue = da.ExecuteDataSet(strSQL)
                If (drValue.Tables(0).Rows.Count = 0) Then
                    Return ""
                End If

                For i = 0 To drValue.Tables(0).Columns.Count - 1
                    value = drValue.Tables(0).Rows(0)(i).ToString
                    If (Not dsvalue.Tables(0).Rows(i)("prefix") Is DBNull.Value) Then
                        value = dsvalue.Tables(0).Rows(i)("prefix") + value
                    End If

                    If (Not dsvalue.Tables(0).Rows(i)("suffix") Is DBNull.Value) Then
                        value = value + dsvalue.Tables(0).Rows(i)("suffix")
                    End If
                    getValue5 = getValue5 + dsvalue.Tables(0).Rows(i)("parameter") + "^" + value + "^"
                Next


            Catch ex As Exception
                ErrorLogging("SFC_LabelPrintOnline_getValue", "", ex.Message & ex.Source)
            End Try
        End Using
    End Function



    Private Function getValue(ByVal dsvalue As DataSet, ByVal labelid As String, ByVal Boxid As String, Optional ByVal productSN As String = "") As String
        Dim da As DataAccess = GetDataAccess()
        getValue = ""
        Dim drValue As DataSet = New DataSet
        Dim i As Integer
        Dim values As String = ""
        Dim value As String = ""
        Dim dr As DataRow() = dsvalue.Tables(0).Select("Type='Unique'")
        Dim strSQL As String = ""

        If (dr.Length = 0) Then
            Return ""
        End If
        Try
            If (dr.Length = 1) Then
                values = dr(0)("Value")
            Else
                For i = 0 To dr.Length - 2
                    values = values + dr(i)("Value") + ","
                Next
                values = values + dr(dr.Length - 1)("Value")
            End If


            ''strSQL = String.Format("Select distinct {0} from V_PackingLabelPrint  where  cartonid='{1}' ", values, Boxid)
            If (productSN = "") Then
                strSQL = String.Format("Select distinct {0} from V_PackingLabelPrint  where  cartonid='{1}' ", values, Boxid)
            Else
                strSQL = String.Format("Select distinct {0} from V_PackingLabelPrint  where  ProductSerialNo='{1}' ", values, productSN)
            End If
            drValue = da.ExecuteDataSet(strSQL)
            If (drValue.Tables(0).Rows.Count = 0) Then
                Return ""
            End If


            Dim drs As DataRow() = dsvalue.Tables(0).Select(String.Format("type='{0}'", "Count"))

            For i = 0 To drs.Length - 1
                dsvalue.Tables(0).Rows.Remove(drs(i))
            Next

            For i = 0 To drValue.Tables(0).Columns.Count - 1
                value = drValue.Tables(0).Rows(0)(i).ToString
                If (Not dsvalue.Tables(0).Rows(i)("prefix") Is DBNull.Value) Then
                    value = dsvalue.Tables(0).Rows(i)("prefix") + value
                End If

                If (Not dsvalue.Tables(0).Rows(i)("suffix") Is DBNull.Value) Then
                    value = value + dsvalue.Tables(0).Rows(i)("suffix")
                End If
                getValue = getValue + dsvalue.Tables(0).Rows(i)("parameter") + "^" + value + "^"
            Next


        Catch ex As Exception
            ErrorLogging("SFC_LabelPrintOnline_getValue", "", ex.Message)
        End Try

    End Function


    Private Function getSNs5(ByVal value As String, ByVal value2 As String, ByVal labelid As String, ByVal boxid As String, ByVal sn2 As ArrayList, Optional ByVal productSN As String = "") As ArrayList
        Using da As DataAccess = GetDataAccess()
            Dim sNs As ArrayList
            Dim times As Integer = 0
            Dim size As Integer = 0
            Dim j As Integer
            Dim drHeader As DataTable
            Dim strSQL As String = ""
            Try

                If (value2 <> "") Then
                    strSQL = String.Format("Select {0} ,{1} from T_shippment where  cartonid= '{2}'", value, value2, boxid)
                    times = 2
                ElseIf (Not String.IsNullOrEmpty(productSN)) Then
                    strSQL = String.Format("Select {0} from T_shippment where  ProductSerialNo = '{1}'", value, productSN)
                Else
                    strSQL = String.Format("Select {0} from T_shippment where  cartonid= '{1}'", value, boxid)
                    times = 1
                End If

                drHeader = da.ExecuteDataTable(strSQL)
                size = drHeader.Rows.Count
                sNs = New ArrayList(times * size)
                For j = 0 To size - 1
                    sNs.Add(drHeader.Rows(j)(0))
                    If (times = 2) Then
                        sn2.Add(drHeader.Rows(j)(1))
                    End If
                Next

                Return sNs
            Catch ex As Exception
                ErrorLogging("SFC_LabelPrintOnline_getSNs", "", ex.Message & ex.Source)
                Return New ArrayList
            End Try
        End Using
    End Function

    Private Function getSNs(ByVal value As String, ByVal value2 As String, ByVal labelid As String, ByVal boxid As String, ByVal sn2 As ArrayList, Optional ByVal productSN As String = "") As ArrayList
        Dim da As DataAccess = GetDataAccess()
        Dim sNs As ArrayList
        Dim times As Integer = 0
        Dim size As Integer = 0
        Dim j As Integer
        Dim drHeader As DataTable
        Dim strSQL As String = ""
        Try

            If (value2 <> "") Then
                strSQL = String.Format("Select {0} ,{1} from T_shippment where  cartonid= '{2}'", value, value2, boxid)
                times = 2
            ElseIf (productSN <> "") Then
                strSQL = String.Format("Select {0} from T_shippment where  ProductSerialNo = '{1}'", value, productSN)
            Else
                strSQL = String.Format("Select {0} from T_shippment where  cartonid= '{1}'", value, boxid)
                times = 1
            End If

            drHeader = da.ExecuteDataTable(strSQL)
            size = drHeader.Rows.Count
            sNs = New ArrayList(times * size)
            For j = 0 To size - 1
                sNs.Add(drHeader.Rows(j)(0))
                If (times = 2) Then
                    sn2.Add(drHeader.Rows(j)(1))
                End If
            Next

            Return sNs
        Catch ex As Exception
            ErrorLogging("SFC_LabelPrintOnline_getSNs", "", ex.Message)
            Return New ArrayList
        End Try

    End Function



    Private Function PrintLabels(ByVal Printer As String, ByVal labelFile As String, ByVal strContent As String, Optional ByVal NoOfLabels As Integer = 1) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Try
                sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", labelFile, Printer, SQLString(strContent))
                da.ExecuteScalar(sql)
                PrintLabels = True
            Catch ex As Exception
                ErrorLogging("SFC-PrintLabels", "", "labelFile: " & labelFile & ", " & ex.Message & ex.Source, "E")
                PrintLabels = ex.Message & ex.Source
            End Try
        End Using
    End Function

    Public Function getLabels(ByVal Model As String, ByVal PCBA As String, ByVal Process As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetLabels '{0}','{1}','{2}'", Model, PCBA, Process)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-getLabels", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function Rework(ByVal ExtSN As String, ByVal DJ As String, ByVal Model As String, ByVal RetestNo As String, ByVal check As Boolean, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPRework '{0}','{1}','{2}',N'{3}','{4}','{5}','{6}',{7},N'{8}'", ExtSN, DJ, Model, ERPLoginData.User, ERPLoginData.OrgCode, ERPLoginData.ProductionLine, RetestNo, check, DStoXML(dsFlow))
                Return da.ExecuteScalar(strSQL).ToString

            Catch ex As Exception
                ErrorLogging("SFC-Rework", "", ex.Message & ex.Source, "E")
                Return "Save Fail!"
            End Try
        End Using
    End Function




    Public Function Rework_New(ByVal dsFlow As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_Rework N'{0}'", DStoXML(dsFlow))
                Return da.ExecuteScalar(strSQL).ToString


            Catch ex As Exception
                ErrorLogging("SFC-Rework", "", ex.Message & ex.Source, "E")
                Return "Save Fail!"
            End Try
        End Using
    End Function

    Public Function TraceLevel(ByVal Model As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetTraceabilityLevel '{0}'", Model)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-TraceLevel", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetPanelSize(ByVal Model As String, ByVal PCBA As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetPanelSize '{0}','{1}'", Model, PCBA)
                Return Convert.ToInt32(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-GetPanelSize", "", ex.Message & ex.Source, "E")
                Return 0
            End Try
        End Using
    End Function

    Public Function IntSNIsValidByPanel(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_IntSNIsValidByPanel '{0}'", IntSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-IntSNIsValidByPanel", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function WIPMatchingByPanel(ByVal DSWIP As DataSet, ByVal PanelSize As Integer) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = "^TDC-156@"
            Try
                strSQL = String.Format("exec sp_WIPMatchingByPanel N'{0}',{1}", DStoXML(DSWIP), PanelSize)
                result = da.ExecuteScalar(strSQL).ToString
                If result = "^TDC-155@" Then
                    strSQL = "exec sp_InsertMatchingRecordByPanelVII N'" + DStoXML(DSWIP) + "','" + PanelSize.ToString + "'"
                    Dim cmdSQL As SqlClient.SqlCommand
                    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceAdditionConnString"))
                    myConn.Open()
                    cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                    result = cmdSQL.ExecuteScalar.ToString()
                End If
            Catch ex As Exception
                ErrorLogging("SFC-sp_WIPMatchingByPanel", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function CheckPanelID(ByVal PanelID As String, ByVal Model As String, ByVal Process As String) As String
        CheckPanelID = "3"

        'Dim cmdPanel As SqlClient.SqlCommand
        Dim cmdProcess As SqlClient.SqlCommand
        Dim adoCon As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceAdditionConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim sSQL As String
        Dim sProcess As String
        sProcess = Process
        If sProcess Is Nothing Then
            sProcess = "Matching1"
        Else
            If sProcess.Trim.Length > 10 Or sProcess.Trim.Length = 0 Then
                sProcess = "Matching1"
            End If
        End If
        Try
            adoCon.Open()
            sSQL = "exec SP_VerifyMatchingVII '" + PanelID + "','" + Model + "','" + sProcess + "'"
            cmdProcess = New SqlClient.SqlCommand(sSQL, adoCon)
            objReader = cmdProcess.ExecuteReader
            Dim sReturnValue As String
            If objReader.Read Then
                sReturnValue = objReader.GetValue(0).ToString.Trim
                If sReturnValue = "1" Then
                    CheckPanelID = "2"
                ElseIf sReturnValue = "2" Then
                    CheckPanelID = "1"
                ElseIf sReturnValue = "3" Then
                    CheckPanelID = "0"
                ElseIf sReturnValue = "4" Then
                    CheckPanelID = "4"
                End If
            Else
                CheckPanelID = "3"
            End If
        Catch ex As Exception
            ErrorLogging("TDC-CheckPanelID", "", "PanelID: " & PanelID & ", " & ex.Message & ex.Source, "E")
            CheckPanelID = "3"
        Finally
            If adoCon.State = ConnectionState.Open Then
                adoCon.Close()
            End If
        End Try
    End Function

    Public Function WIPRework(ByVal IntSN As String, ByVal ERPLoginData As ERPLogin, ByVal dsFlow As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WIPReworkForWIP '{0}',N'{1}',N'{2}'", IntSN, ERPLoginData.User, DStoXML(dsFlow))
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPRework", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function ReadMIData(ByVal model As String, ByVal pcba As String, ByVal process As String, ByVal status As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ItemData As DataTable
            Dim myDataColumn As DataColumn
            Dim sqlstr As String = ""
            ReadMIData = New DataSet()

            ItemData = New Data.DataTable("ItemData")
            myDataColumn = New Data.DataColumn("MIId", System.Type.GetType("System.Int16"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Model", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("PCBA", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Process", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ValidFrom", System.Type.GetType("System.DateTime"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ValidTo", System.Type.GetType("System.DateTime"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MIStatus", System.Type.GetType("System.Int16"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("CreatedOn", System.Type.GetType("System.DateTime"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("CreatedBy", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("LastChangedOn", System.Type.GetType("System.DateTime"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("LastChangedBy", System.Type.GetType("System.String"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MIFile", System.Type.GetType("System.Object"))
            ItemData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("RowState", System.Type.GetType("System.Int16"))
            ItemData.Columns.Add(myDataColumn)
            ReadMIData.Tables.Add(ItemData)


            Dim ErrorTable As DataTable
            Dim ErrorRow As Data.DataRow
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            ReadMIData.Tables.Add(ErrorTable)


            If (model.Length > 0) Then
                sqlstr += " and model like '" + model.Trim + "%'"
            End If

            If (pcba.Length > 0) Then
                sqlstr += " and pcba like '" + pcba.Trim + "%'"
            End If

            If (process.Length > 0) Then
                sqlstr += " and process like '" + process.Trim + "%'"
            End If

            If (status.ToUpper = "VALID") Then
                sqlstr += " and status =1"
            ElseIf (status.ToUpper = "INVALID") Then
                sqlstr += " and status =0"
            End If

            sqlstr = String.Format("select * ,0 as rowstate from T_MIItems where 0=0  {0}", sqlstr)
            Try
                ReadMIData = da.ExecuteDataSet(sqlstr, "ItemData")

                If ReadMIData.Tables("ItemData").Rows.Count = 0 Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrorMsg") = "No record found!"
                    ErrorTable.Rows.Add(ErrorRow)
                End If
            Catch ex As Exception
                ErrorLogging("SFC-ReadMIData", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function
    'Public Function GetOrderInfoFromOracle(ByVal OrderNo As String) As DataSet

    '    Using da As DataAccess = GetDataAccess()
    '        Dim oda As OracleDataAdapter = da.Oda_Sele()
    '        Dim ds As New DataSet()
    '        Try
    '            ds.Tables.Add("DJInfo")
    '            oda.SelectCommand.CommandType = CommandType.StoredProcedure
    '            oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_dj"

    '            oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
    '            oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = OrderNo
    '            oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = "580"

    '            oda.SelectCommand.Connection.Open()
    '            oda.Fill(ds, "DJInfo")
    '            oda.SelectCommand.Connection.Close()

    '            If ds.Tables("DJInfo").Rows.Count = 0 Then
    '                ds = New DataSet
    '                ds = GetOrderInfoFromETRACE(OrderNo)
    '            End If
    '        Catch oe As Exception
    '            ErrorLogging("TDCService-GetOrderInfoFromERP", "", "DJ: " & OrderNo & ", " & oe.Message & oe.Source, "E")

    '            ds = GetOrderInfoFromETRACE(OrderNo)

    '        Finally
    '            If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
    '        End Try
    '        Return ds
    '    End Using

    'End Function
    Public Function GetOrderInfoFromOracle(ByVal OrderNo As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                ds.Tables.Add("DJInfo")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_dj"

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = OrderNo
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = "580"

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "DJInfo")
                oda.SelectCommand.Connection.Close()

                If ds.Tables("DJInfo").Rows.Count = 0 Then
                    ds.Clear()
                    ds = New DataSet()

                    If Not ds.Tables.Contains("DJInfo") Then
                        ds.Tables.Add("DJInfo")
                    End If
                    oda.SelectCommand.Parameters.Clear()
                    oda.SelectCommand.CommandType = CommandType.StoredProcedure
                    oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_dj"

                    oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                    oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = OrderNo
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = "680"

                    oda.SelectCommand.Connection.Open()
                    oda.Fill(ds, "DJInfo")
                    oda.SelectCommand.Connection.Close()
                    If ds.Tables("DJInfo").Rows.Count = 0 Then
                        ds = New DataSet
                        ds = GetOrderInfoFromETRACE(OrderNo)
                    Else
                        ds.Tables("DJInfo").Columns.Add(New Data.DataColumn("OrgCode", System.Type.GetType("System.String")))
                        ds.Tables("DJInfo").Rows(0)("OrgCode") = "680"
                    End If
                Else
                    ds.Tables("DJInfo").Columns.Add(New Data.DataColumn("OrgCode", System.Type.GetType("System.String")))
                    ds.Tables("DJInfo").Rows(0)("OrgCode") = "580"
                End If
            Catch oe As Exception
                ErrorLogging("TDCService-GetOrderInfoFromOracle", "", "DJ: " & OrderNo & ", " & oe.Message & oe.Source, "E")

                ds = GetOrderInfoFromETRACE(OrderNo)

            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
            Return ds
        End Using
    End Function
    Public Function getMIFileData(ByVal model As String, ByVal PCBA As String, ByVal process As String) As Object
        Dim sqlstr As String
        Using da As DataAccess = GetDataAccess()
            sqlstr = String.Format("sp_getMIFileData '{0}','{1}','{2}'", model, PCBA, process)
            Try
                getMIFileData = da.ExecuteScalar(sqlstr)
            Catch ex As Exception
                ErrorLogging("SFC-getMIFileData", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function GetOrderInfoFromETRACE(ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                Dim ds As New DataSet()
                sqlstr = String.Format("exec sp_GetOrderInfoFromETRACE '{0}'", OrderNo)
                ds = da.ExecuteDataSet(sqlstr, "DJInfo")
                Return ds
            Catch ex As Exception
                Throw ex
                ErrorLogging("TDCService-GetOrderInfoFromETRACE", "", "DJ: " & OrderNo & ", " & ex.Message & ex.Source)
            End Try
        End Using
    End Function
    Public Function SaveMIRecord(ByVal MIRecords As DataSet, ByVal username As String) As Boolean
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Sqlconn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim dr As DataRow
        SaveMIRecord = True
        Try
            sqlCommand = New SqlClient.SqlCommand("sp_UpdateEMI", Sqlconn)
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Parameters.Add("@ValidFrom", SqlDbType.SmallDateTime, 4, "ValidFrom")
            sqlCommand.Parameters.Add("@ValidTo", SqlDbType.SmallDateTime, 4, "ValidTo")
            sqlCommand.Parameters.Add("@MIFile", SqlDbType.Image, 5000000, "MIFile")
            sqlCommand.Parameters.Add("@Status", SqlDbType.Int, 1, "Status")
            sqlCommand.Parameters.Add("@username", SqlDbType.NVarChar, 50, "username")
            sqlCommand.Parameters.Add("@MIID", SqlDbType.Int, 20, "MIID")
            sqlCommand.Parameters.Add("@Model", SqlDbType.VarChar, 50, "Model")
            sqlCommand.Parameters.Add("@PCBA", SqlDbType.VarChar, 50, "PCBA")
            sqlCommand.Parameters.Add("@Process", SqlDbType.VarChar, 50, "Process")
            sqlCommand.Parameters.Add("@rowstate", SqlDbType.Int, 2, "rowstate")


            Sqlconn.Open()
            Dim i As Integer
            For i = 0 To MIRecords.Tables(0).Rows.Count - 1
                dr = MIRecords.Tables(0).Rows(i)
                If (dr("rowState") = DataRowState.Modified Or dr("rowState") = DataRowState.Added) Then
                    If (dr("MIID").ToString.Length = 0) Then
                        dr("MIID") = 0
                    End If
                    sqlCommand.Parameters("@Model").Value = dr("Model").ToString.ToUpper()
                    sqlCommand.Parameters("@PCBA").Value = dr("PCBA").ToString.ToUpper()
                    sqlCommand.Parameters("@Process").Value = dr("Process").ToString.ToUpper()
                    If (MIRecords.Tables(0).Rows(i)("Status") Is DBNull.Value) Then
                        sqlCommand.Parameters("@Status").Value = 0
                    Else
                        sqlCommand.Parameters("@Status").Value = dr("Status")
                    End If
                    sqlCommand.Parameters("@MIFile").Value = dr("MIFile")
                    sqlCommand.Parameters("@ValidFrom").Value = dr("ValidFrom")
                    sqlCommand.Parameters("@ValidTo").Value = dr("ValidTo")
                    sqlCommand.Parameters("@MIID").Value = dr("MIID")
                    sqlCommand.Parameters("@rowstate").Value = dr("rowstate")
                    sqlCommand.Parameters("@username").Value = username
                    sqlCommand.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            ErrorLogging("SFC-SaveMIRecord", username, ex.Message & ex.Source, "E")
            SaveMIRecord = False
        Finally
            If Sqlconn.State = ConnectionState.Open Then Sqlconn.Close()
        End Try
    End Function

    Public Function GetConfig(ByVal eTraceModule As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetConfig '{0}'", eTraceModule)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-GetConfig", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetModelByIntSN(ByVal intSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetModelByIntSN '{0}'", intSN)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-GetModelByIntSN", IIf(intSN Is Nothing, "Nothing", intSN), ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetPCBAByIntSN(ByVal intSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetPCBAByIntSN '{0}'", intSN)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-GetPCBAByIntSN", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function Getwipflowdata(ByVal intSN As String, ByVal ProcessName As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetWIPflowData '{0}','{1}'", intSN, ProcessName)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-Getwipflowdata", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetMaxTestRound(ByVal intSN As String, ByVal ProcessName As String) As Integer
        GetMaxTestRound = -1
        Try
            Dim dtfl As DataTable = Getwipflowdata(intSN, ProcessName).Tables(0)
            If dtfl.Rows.Count > 0 Then
                GetMaxTestRound = Convert.ToInt32(dtfl.Rows(0).Item("MaxTestRound"))
            End If
        Catch ex As Exception
            ErrorLogging("SFC-GetMaxTestRound", "", ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function GetMaxFailure(ByVal intSN As String, ByVal ProcessName As String) As Integer
        GetMaxFailure = -1
        Try
            Dim dtfl As DataTable = Getwipflowdata(intSN, ProcessName).Tables(0)
            If dtfl.Rows.Count > 0 Then
                GetMaxFailure = Convert.ToInt32(dtfl.Rows(0).Item("MaxFailure"))
            End If
        Catch ex As Exception
            ErrorLogging("SFC-GetMaxFailure", "", ex.Message & ex.Source, "E")
        End Try
    End Function


    Public Function GetLastTestResult(ByVal intSN As String, ByVal ProcessName As String) As String
        GetLastTestResult = ""
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetLastTestResult '{0}','{1}'", intSN, ProcessName)
                GetLastTestResult = da.ExecuteScalar(strSQL)
                If (FixNull(GetLastTestResult) = "") Then
                    GetLastTestResult = ""
                Else
                    GetLastTestResult = GetLastTestResult.ToString
                End If
            Catch ex As Exception
                ErrorLogging("SFC-GetLastTestResult", "", ex.Message & ex.Source, "E")

            End Try
        End Using
    End Function

    ' Temp Funvtion just for TE to test
    Public Function CleanTestResult(ByVal intSN As String, ByVal ProcessName As String) As Boolean
        CleanTestResult = False
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("UPDATE T_WIPFlow SET LastResult = '', status='' WHERE (Process = '{1}') AND (WIPID = (SELECT WIPID FROM T_WIPHeader WHERE (IntSN = '{0}'))) ", intSN, ProcessName)
                da.ExecuteNonQuery(strSQL)
                CleanTestResult = True
            Catch ex As Exception
                ErrorLogging("SFC-CleanTestResult", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetPropertiesName(ByVal intSN As String, ByVal processname As String) As String
        GetPropertiesName = ""
        Dim PropertiesName As DataTable = New DataTable
        Dim i As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetPropertiesNames '{0}','{1}'", intSN, processname)
                PropertiesName = da.ExecuteDataTable(strSQL)
                If (PropertiesName.Rows.Count = 0) Then
                    GetPropertiesName = ""
                Else
                    For i = 0 To PropertiesName.Rows.Count - 1
                        GetPropertiesName = GetPropertiesName + PropertiesName.Rows(i).Item(0) + ","
                    Next
                    GetPropertiesName = GetPropertiesName.Substring(0, GetPropertiesName.Length - 1)
                End If
            Catch ex As Exception
                ErrorLogging("SFC-GetPropertiesName", "", ex.Message & ex.Source, "E")

            End Try
        End Using
    End Function

    Public Function WIPInOQA(ByVal ExtSN As String, ByVal OperatorName As String, ByVal OrgCode As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String
            Try
                strSQL = String.Format("exec sp_WIPInOQA '{0}','{1}','{2}'", ExtSN, OperatorName, OrgCode)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                result = "OQA WIP IN FAIL"
                ErrorLogging("SFC-ProdQty", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function CheckExtSNSameIntSNByModel(ByVal Model As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_CheckExtSNSameIntSNByModel '{0}'", Model)
                Return Convert.ToBoolean(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC-CheckExtSNSameIntSNByModel", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function CopyCardInfo() As Boolean
        Dim i As Integer

        Dim dsCardInfo As DataSet = New DataSet("dsCardInfo")
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn01 As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("HRISDBConnString"))
        Try
            'Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select Card, Number,ChineseName from Employees where IncumbencyStatus like '0%' "), myConn01)
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select AccessCardID,EmployeeID,Name,Dept,Site,CertifiedCourse from eTrace_AccessCard "), myConn01)
            da.Fill(dsCardInfo)
        Catch ex As Exception
            ErrorLogging("SFC-CopyCardInfo", "", ex.Message & ex.Source, "E")
        End Try

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        myConn.Open()
        Dim myTrans As SqlClient.SqlTransaction = myConn.BeginTransaction()
        Dim myCommand As SqlClient.SqlCommand = myConn.CreateCommand()
        myCommand.Transaction = myTrans
        Try
            myCommand.CommandText = "delete from T_AccessCardBackup"
            myCommand.ExecuteNonQuery()
            'myCommand.CommandText = "Insert into T_AccessCardBackup(AccessCardID,EmployeeID,Name) select AccessCardID,EmployeeID,Name from T_AccessCard"
            myCommand.CommandText = "Insert into T_AccessCardBackup(AccessCardID,EmployeeID,Name,Dept,CertifiedCourse) select AccessCardID,EmployeeID,Name,Dept,CertifiedCourse from T_AccessCard"
            myCommand.ExecuteNonQuery()
            myCommand.CommandText = "delete from T_AccessCard"
            myCommand.ExecuteNonQuery()
            For i = 0 To dsCardInfo.Tables(0).Rows.Count - 1
                'myCommand.CommandText = String.Format("insert into T_AccessCard(AccessCardID,EmployeeID,Name) values('{0}','{1}',N'{2}')", dsCardInfo.Tables(0).Rows(i)(0), dsCardInfo.Tables(0).Rows(i)(1), dsCardInfo.Tables(0).Rows(i)(2))
                myCommand.CommandText = String.Format("insert into T_AccessCard(AccessCardID,EmployeeID,Name,Dept,Site,CertifiedCourse) values('{0}','{1}',N'{2}','{3}','{4}','{5}')", dsCardInfo.Tables(0).Rows(i)(0), dsCardInfo.Tables(0).Rows(i)(1), dsCardInfo.Tables(0).Rows(i)(2), dsCardInfo.Tables(0).Rows(i)(3), dsCardInfo.Tables(0).Rows(i)(4), dsCardInfo.Tables(0).Rows(i)(5))
                myCommand.ExecuteNonQuery()
            Next
            myTrans.Commit()
            myConn.Close()
            Return True
        Catch ex As Exception
            Try
                myTrans.Rollback()
                ErrorLogging("SFC-CopyCardInfo", "", ex.Message & ex.Source, "E")
            Catch ex1 As Exception
                ErrorLogging("SFC-CopyCardInfo-rollback", "", ex1.Message & ex1.Source, "E")
            End Try
            Return False
        End Try
    End Function

    Public Function CopyCardInfoZS() As Boolean
        Dim i As Integer

        Dim dsCardInfo As DataSet = New DataSet("dsCardInfo")
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn01 As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("HRISDBConnString"))
        Try
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select Cust2668, Number,ChineseName,Cust2260 from Employees where IncumbencyStatus like '0%' and not Cust2668 is NULL and Cust2668<>'' "), myConn01)
            'Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("Select AccessCardID,EmployeeID,Name,Dept,Site,CertifiedCourse from eTrace_AccessCard "), myConn01)
            da.Fill(dsCardInfo)
        Catch ex As Exception
            ErrorLogging("SFC-CopyCardInfo", "", ex.Message & ex.Source, "E")
        End Try

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        myConn.Open()
        Dim myTrans As SqlClient.SqlTransaction = myConn.BeginTransaction()
        Dim myCommand As SqlClient.SqlCommand = myConn.CreateCommand()
        myCommand.Transaction = myTrans
        Try
            myCommand.CommandText = "delete from T_AccessCardBackup"
            myCommand.ExecuteNonQuery()
            myCommand.CommandText = "Insert into T_AccessCardBackup(AccessCardID,EmployeeID,Name,Dept) select AccessCardID,EmployeeID,Name,Dept from T_AccessCard"
            'myCommand.CommandText = "Insert into T_AccessCardBackup(AccessCardID,EmployeeID,Name,Dept,CertifiedCourse) select AccessCardID,EmployeeID,Name,Dept,CertifiedCourse from T_AccessCard"
            myCommand.ExecuteNonQuery()
            myCommand.CommandText = "delete from T_AccessCard"
            myCommand.ExecuteNonQuery()
            For i = 0 To dsCardInfo.Tables(0).Rows.Count - 1
                myCommand.CommandText = String.Format("insert into T_AccessCard(AccessCardID,EmployeeID,Name,Dept) values('{0}','{1}',N'{2}',N'{3}')", dsCardInfo.Tables(0).Rows(i)(0), dsCardInfo.Tables(0).Rows(i)(1), dsCardInfo.Tables(0).Rows(i)(2), dsCardInfo.Tables(0).Rows(i)(3))
                'myCommand.CommandText = String.Format("insert into T_AccessCard(AccessCardID,EmployeeID,Name,Dept,Site,CertifiedCourse) values('{0}','{1}',N'{2}','{3}','{4}','{5}')", dsCardInfo.Tables(0).Rows(i)(0), dsCardInfo.Tables(0).Rows(i)(1), dsCardInfo.Tables(0).Rows(i)(2), dsCardInfo.Tables(0).Rows(i)(3), dsCardInfo.Tables(0).Rows(i)(4), dsCardInfo.Tables(0).Rows(i)(5))
                myCommand.ExecuteNonQuery()
            Next
            myTrans.Commit()
            myConn.Close()
            Return True
        Catch ex As Exception
            Try
                myTrans.Rollback()
                ErrorLogging("SFC-CopyCardInfoZS", "", ex.Message & ex.Source, "E")
            Catch ex1 As Exception
                ErrorLogging("SFC-CopyCardInfoZS-rollback", "", ex1.Message & ex1.Source, "E")
            End Try
            Return False
        End Try
    End Function

    Public Function Fixture_Register(ByVal Type As String, ByVal FixtureID As String, ByVal MaxUse As Integer, ByVal Description As String, ByVal User As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FixtureRegister '{0}','{1}','{2}',N'{3}',N'{4}'", Type, FixtureID, MaxUse, Description, User)
                Return da.ExecuteScalar(StrSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-FixtureRegister", "", ex.Message & ex.Source, "E")
                Return "3"
            End Try
        End Using
    End Function

    Public Function Fixture_Type(ByVal Type As String, ByVal Slot As Integer, ByVal Warning As Integer, ByVal Maintenance As Integer, ByVal Repair As Integer, ByVal DefaultMax As Integer, ByVal Description As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FixtureType '{0}','{1}','{2}','{3}','{4}','{5}',N'{6}',N'{7}'", Type, Slot, Warning, Maintenance, Repair, DefaultMax, Description, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-FixtureType", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function Fixture_Maintain(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FixtureMaintain '{0}',N'{1}',N'{2}','Maintenance'", FixtureID, Description, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-FixtureMaintain", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function Fixture_Repair(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FixtureMaintain '{0}',N'{1}',N'{2}','Repair'", FixtureID, Description, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-FixtureRepair", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function Fixture_InActive(ByVal FixtureID As String, ByVal Description As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FixtureInActive '{0}',N'{1}',N'{2}'", FixtureID, Description, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-FixtureInActive", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function Fixture_RegisterView(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureRegisterView '{0}','1'", FixtureID)
                dsTmp = da.ExecuteDataSet(StrSQL)
                dsTmp.Tables(0).TableName = "tblFixture"
                Dim tt As String
                tt = dsTmp.GetXml()
                StrSQL = String.Format("exec SP_FixtureRegisterView '{0}','0'", FixtureID)
                dsTmp.Tables.Add(da.ExecuteDataTable(StrSQL))
                dsTmp.Tables(1).TableName = "tblFixtureDetail"
                Return dsTmp.GetXml
            Catch ex As Exception
                ErrorLogging("SFC-FixtureRegisterReview", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_InActiveLog() As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureInActiveLog ")
                dsTmp = da.ExecuteDataSet(StrSQL)
                dsTmp.Tables(0).TableName = "tblFixture"
                Return dsTmp.GetXml
            Catch ex As Exception
                ErrorLogging("SFC-FixtureInActiveLog", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_InActiveLogByFixture(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureInActiveLogByFixture '{0}'", FixtureID)
                dsTmp = da.ExecuteDataSet(StrSQL)
                dsTmp.Tables(0).TableName = "tblFixture"
                Return dsTmp.GetXml
            Catch ex As Exception
                ErrorLogging("SFC-FixtureInActiveLogByFixture", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_MaintainLog(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureMaintainLog '{0}','Maintenance'", FixtureID)
                Return da.ExecuteDataSet(StrSQL).GetXml()
            Catch ex As Exception
                ErrorLogging("SFC-FixtureMaintainLog", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_RepairLog(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureMaintainLog '{0}','Repair'", FixtureID)
                Return da.ExecuteDataSet(StrSQL).GetXml()
            Catch ex As Exception
                ErrorLogging("SFC-FixtureRepairLog", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_TypeView() As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureTypeView ")
                Return da.ExecuteDataSet(StrSQL).GetXml
            Catch ex As Exception
                ErrorLogging("SFC-FixtureTypeView", "", ex.Message & ex.Source, "E")
                Return dsTmp.GetXml
            End Try
        End Using
    End Function

    Public Function Fixture_Update(ByVal FixtureID As String, ByVal Type As String, ByVal MaxUse As Integer, ByVal Description As String, ByVal User As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Dim dsTmp As New DataSet
            Try
                StrSQL = String.Format("exec SP_FixtureUpdate '{0}','{1}',{2},N'{3}',N'{4}'", FixtureID, Type, MaxUse, Description, User)
                Return da.ExecuteScalar(StrSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-FixtureUpdate", "", ex.Message & ex.Source, "E")
                Return "0"
            End Try
        End Using
    End Function

    Public Function Employee_Login(ByVal AccessCardID As String) As String
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        'Dim dt As DataTable
        Try
            'Dim daTmp As New SqlClient.SqlDataAdapter
            'myConn.Open()
            strSQL = String.Format("select AccessCardID,EmployeeID,Name from T_AccessCard where AccessCardID='" + AccessCardID + "'")
            'daTmp.SelectCommand = New SqlClient.SqlCommand(strSQL, myConn)
            'daTmp.Fill(dsTmp, "Employee")
            'dt = da.ExecuteDataTable(strSQL, CommandType.Text)
            'dt.TableName = "Employee"
            'dsTmp.Tables.Add(dt)
            dsTmp = da.ExecuteDataSet(strSQL)
            If Not (dsTmp Is Nothing) Then
                If Not (dsTmp.Tables Is Nothing) Then
                    dsTmp.Tables(0).TableName = "Employee"
                End If
            End If
            Return dsTmp.GetXml
        Catch ex As Exception
            ErrorLogging("SFC-Employee_Login", "", ex.Message & ex.Source, "E")
            Return dsTmp.GetXml
        End Try
    End Function

    Public Function Employee_Certify(ByVal AccessCardID As String, ByVal CourseCode As String) As String
        'Dim cmdSQL As SqlClient.SqlCommand
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Try
            'Dim daTmp As New SqlClient.SqlDataAdapter
            'myConn.Open()
            strSQL = String.Format("select AccessCardID,EmployeeID,Name from T_AccessCard where AccessCardID='" + AccessCardID + "' and CertifiedCourse like '%" + CourseCode + "%'")
            'daTmp.SelectCommand = New SqlClient.SqlCommand(strSQL, myConn)
            'daTmp.Fill(dsTmp, "Employee")
            dsTmp = da.ExecuteDataSet(strSQL)
            If Not (dsTmp Is Nothing) Then
                If Not (dsTmp.Tables Is Nothing) Then
                    dsTmp.Tables(0).TableName = "Employee"
                End If
            End If
            Return dsTmp.GetXml
        Catch ex As Exception
            ErrorLogging("SFC-Employee_Certify", "", ex.Message & ex.Source, "E")
            Return dsTmp.GetXml
        End Try
    End Function

    Public Function AutoStopLine(ByVal TestData As String, ByVal Type As String) As String
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim da As DataAccess = GetDataAccess()
        Try
            Dim dsTmp As New DataSet
            'Dim daTmp As New SqlClient.SqlDataAdapter
            'myConn.Open()
            strSQL = String.Format("exec eTraceAddition.dbo.sp_AutoStopLine '{0}',{1}", TestData, Type)
            'daTmp.SelectCommand = New SqlClient.SqlCommand(strSQL, myConn)
            'daTmp.Fill(dsTmp, "tblResult")
            Dim dt As DataTable
            dt = da.ExecuteDataTable(strSQL, CommandType.StoredProcedure)
            'If dsTmp Is Nothing Then
            '    Return ""
            'End If
            'If dsTmp.Tables Is Nothing Then
            '    Return ""
            'End If
            'If dsTmp.Tables.Count < 1 Then
            '    Return ""
            'End If
            If dt Is Nothing Then
                Return ""
            End If
            'If dsTmp.Tables("tblResult").Rows.Count < 1 Then
            '    Return ""
            'End If
            If dt.Rows.Count < 1 Then
                Return ""
            End If
            'If Type = 1 Then
            '    Return dsTmp.GetXml
            'Else
            '    Return dsTmp.Tables("tblResult").Rows(0)(0).ToString
            'End If
            If Type = 1 Then
                dt.TableName = "tblResult"
                dsTmp.Tables.Add(dt)
                Return dsTmp.GetXml
            Else
                Return dt.Rows(0)(0).ToString
            End If
        Catch ex As Exception
            ErrorLogging("TDC-AutoStopLine", "", ex.Message & ex.Source, "E")
            Return ""
        End Try
    End Function

    Public Function LDTrigger_GetRepairData(ByVal Model As String, ByVal Line As String, ByVal Station As String) As String
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Try
            strSQL = String.Format("exec SP_LDTriggerGetRepairData '" + Model + "','" + Line + "','" + Station + "'")
            dsTmp = da.ExecuteDataSet(strSQL)
            dsTmp.Tables(0).TableName = "tblRepair"
            Return dsTmp.GetXml.ToString
        Catch ex As Exception
            ErrorLogging("SFC-LDTrigger_GetRepairData", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Loading_CheckModel(ByVal Model As String) As DataSet
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Try
            strSQL = String.Format("exec ST_LoadingCheckModel '" + Model + "'")
            dsTmp = da.ExecuteDataSet(strSQL)
            If Not (dsTmp Is Nothing) Then
                If Not (dsTmp.Tables Is Nothing) Then
                    dsTmp.Tables(0).TableName = "tblLoading"
                End If
            End If
            Return dsTmp

        Catch ex As Exception
            ErrorLogging("SFC-Loading_CheckModel", "", ex.Message & ex.Source, "E")
            Return dsTmp
        End Try
    End Function

    Public Function Laser_VerifyUnit(ByVal Model As String, ByVal IntSN As String, ByVal Computer As String, ByVal User As String) As String
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim cmdSQL As SqlClient.SqlCommand
        Dim conSQL As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceAdditionConnString"))
        Try
            Dim sTmp As String
            conSQL.Open()
            strSQL = String.Format("exec SP_VerifyUnitStepI '" + IntSN + "'")
            cmdSQL = New SqlClient.SqlCommand(strSQL, conSQL)
            sTmp = cmdSQL.ExecuteScalar.ToString
            If sTmp <> "4" Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitStepII '" + Model + "', '" + IntSN + "'")
            sTmp = da.ExecuteScalar(strSQL).ToString
            If sTmp <> 6 Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitStepIII '" + Model + "','" + IntSN + "','" + Computer + "','" + User + "'")
            cmdSQL = New SqlClient.SqlCommand(strSQL, conSQL)
            sTmp = cmdSQL.ExecuteScalar.ToString
            If (sTmp = "6") Or (sTmp = "7") Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitStepIV '" + IntSN + "','" + sTmp + "','" + User + "'")
            sTmp = da.ExecuteScalar(strSQL).ToString
            Return sTmp
        Catch ex As Exception
            ErrorLogging("SFC-Laser_VerifyUnit", "", ex.Message & ex.Source, "E")
            Return "7"
        End Try
    End Function

    Public Function Laser_VerifyUnitTest(ByVal Model As String, ByVal IntSN As String, ByVal Computer As String, ByVal User As String) As String
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim cmdSQL As SqlClient.SqlCommand
        Dim conSQL As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceAdditionConnString"))
        Try
            Dim sTmp As String
            conSQL.Open()
            strSQL = String.Format("exec SP_VerifyUnitTestStepI '" + IntSN + "'")
            cmdSQL = New SqlClient.SqlCommand(strSQL, conSQL)
            sTmp = cmdSQL.ExecuteScalar.ToString
            If sTmp <> "4" Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitTestStepII '" + Model + "', '" + IntSN + "'")
            sTmp = da.ExecuteScalar(strSQL)
            If sTmp <> 6 Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitTestStepIII '" + Model + "','" + IntSN + "','" + Computer + "','" + User + "'")
            cmdSQL = New SqlClient.SqlCommand(strSQL, conSQL)
            sTmp = cmdSQL.ExecuteScalar.ToString
            If (sTmp = "6") Or (sTmp = "7") Then
                Return sTmp
                Exit Function
            End If
            strSQL = String.Format("exec SP_VerifyUnitTestStepIV '" + IntSN + "','" + sTmp + "','" + User + "'")
            sTmp = da.ExecuteScalar(strSQL)
            Return sTmp
        Catch ex As Exception
            ErrorLogging("SFC-Laser_VerifyUnitTest", "", ex.Message & ex.Source, "E")
            Return "7"
        End Try
    End Function

    Public Function ATE_FixtureVerify(ByVal FixtureID As String, ByVal Type As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEFixtureVerify '{0}','{1}'", FixtureID, Type)
                Return da.ExecuteScalar(StrSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-ATEFixtureVerify", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function ATE_CreateRelation(ByVal FixtureID As String, ByVal Slot As Integer, ByVal Model As String, ByVal IntSN As String, ByVal Process As String, ByVal User As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                Dim sTmp As String
                Dim cmdSQL As SqlClient.SqlCommand
                Dim conSQL As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceAdditionConnString"))
                StrSQL = "exec sp_ATEVerifyLaser '" + IntSN + "'"
                conSQL.Open()
                cmdSQL = New SqlClient.SqlCommand(StrSQL, conSQL)
                sTmp = cmdSQL.ExecuteScalar.ToString()
                If sTmp.Trim <> "0" Then
                    Return sTmp.Trim
                    Exit Function
                End If
                StrSQL = String.Format("exec SP_ATECreateRelation '{0}','{1}','{2}','{3}','{4}',N'{5}'", FixtureID, Slot, Model, IntSN, Process, User)
                sTmp = da.ExecuteScalar(StrSQL).ToString
                If sTmp = 7 Then
                    StrSQL = "update T_CrossReference set Remark='" + FixtureID + ("00" + Slot.ToString).PadRight(3) + "' where IntSerialNo='" + IntSN + "'"
                    cmdSQL = New SqlClient.SqlCommand(StrSQL, conSQL)
                    Try
                        cmdSQL.ExecuteNonQuery()
                    Catch ex As Exception
                        ErrorLogging("SFC-ATECreateRelation", "", "Update T_CrossReference:" & ex.Message & ex.Source, "E")
                        Return "0"
                    End Try
                    Return "7"
                End If
                Return sTmp
            Catch ex As Exception
                ErrorLogging("SFC-ATECreateRelation", "", ex.Message & ex.Source, "E")
                Return "0"
            End Try
        End Using
    End Function

    Public Function ATE_FixtureSign(ByVal FixtureID As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEFixtureSign '{0}',N'{1}'", FixtureID, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-ATEFixtureSign", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function ATE_ReturnSNbyFixture(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEReturnSN '{0}',0,'Fixture'", FixtureID)
                Return da.ExecuteDataSet(StrSQL).GetXml
            Catch ex As Exception
                ErrorLogging("SFC-ATEReturnSNbyFixture", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function ATE_ReturnSNbySlot(ByVal FixtureID As String, ByVal Slot As Integer) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEReturnSN '{0}',{1},'Slot'", FixtureID, Slot)
                Return da.ExecuteDataSet(StrSQL).GetXml
            Catch ex As Exception
                ErrorLogging("SFC-ATEReturnSNbySlot", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function ATE_ReleaseRelationbySlot(ByVal FixtureID As String, ByVal Slot As Integer, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEReleaseRelation '{0}',{1},'Slot',N'{2}'", FixtureID, Slot, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-ATEReleaseRelationbySlot", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function ATE_ReleaseRelationbyFixture(ByVal FixtureID As String, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEReleaseRelation '{0}',0,'Fixture',N'{1}'", FixtureID, User)
                Return Convert.ToBoolean(da.ExecuteScalar(StrSQL))
            Catch ex As Exception
                ErrorLogging("SFC-ATEReleaseRelationbyFixture", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function ATE_IntSlotReview(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_ATEIntSlotReview '{0}'", IntSN)
                Return da.ExecuteDataSet(StrSQL).GetXml
            Catch ex As Exception
                ErrorLogging("SFC-ATEIntSlotReview", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function Function_ProcessVerify(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_FunctionProcessVerify '{0}'", IntSN)
                Return da.ExecuteDataSet(StrSQL).GetXml
            Catch ex As Exception
                ErrorLogging("SFC-ATEIntSlotReview", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function Depanel_VerifyMatching1(ByVal IntSN As String) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceAdditionConnString"))
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Try
            Dim daTmp As New SqlClient.SqlDataAdapter
            myConn.Open()
            strSQL = String.Format("exec SP_DepanelVerifyMatching1 '" + IntSN + "'")
            daTmp.SelectCommand = New SqlClient.SqlCommand(strSQL, myConn)
            daTmp.Fill(dsTmp, "tblSN")
            Return dsTmp
        Catch ex As Exception
            ErrorLogging("SFC-Depanel_VerifyMatching1", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Depanel_VerifyLastTest(ByVal IntSN As DataSet) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Try
            Dim daTmp As New SqlClient.SqlDataAdapter
            myConn.Open()
            strSQL = String.Format("exec SP_DepanelVerifyLastTest N'" + IntSN.GetXml + "','OVP'")
            daTmp.SelectCommand = New SqlClient.SqlCommand(strSQL, myConn)
            daTmp.Fill(dsTmp, "tblSN")
            Return dsTmp
        Catch ex As Exception
            ErrorLogging("SFC-Depanel_VerifyLastTest", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Test_Reflow(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_TestReflow '{0}'", IntSN)
                Return da.ExecuteDataSet(StrSQL)
            Catch ex As Exception
                ErrorLogging("SFC-TestReflow", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function Test_ItemData(ByVal IntSN As String, ByVal Process As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_TestItemData '{0}','{1}'", IntSN, Process)
                Return da.ExecuteDataSet(StrSQL)
            Catch ex As Exception
                ErrorLogging("SFC-TestItemData", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function Test_ReflowProcess(ByVal IntSN As String, ByVal Process As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                StrSQL = String.Format("exec SP_TestReflowProcess '{0}','{1}'", IntSN, Process)
                Return da.ExecuteScalar(StrSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-TestReflowProcess", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetSystemGMTDateTime(Optional ByVal TimeZone As String = "GMT") As String
        Try
            If TimeZone.ToUpper <> "Local".ToUpper Then
                GetSystemGMTDateTime = Format(Now(), "MM-dd-yyyy HH:mm:ss")
            Else
                GetSystemGMTDateTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
            End If
        Catch ex As Exception
            If TimeZone.ToUpper <> "Local".ToUpper Then
                GetSystemGMTDateTime = Format(Now(), "MM-dd-yyyy HH:mm:ss")
            Else
                GetSystemGMTDateTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
            End If
            ErrorLogging("SFC-GetSystemGMTDateTime", "", ex.Message & ex.Source, "E")
        Finally

        End Try
    End Function

    Public Function WIP_UpdateStatus(ByVal IntSN As String, ByVal Process As String, ByVal OperatorName As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_UpdateStatus '{0}','{1}','{2}'", IntSN, Process, OperatorName)
                Return da.ExecuteScalar(strSQL).ToString()
            Catch ex As Exception
                ErrorLogging("SFC-WIPUpdateStatus", OperatorName, ex.Message & ex.Source, "E")
                Return "FAIL"
            End Try
        End Using
    End Function

    Public Function Get_WIPFGSN(ByVal IntSN As String) As DataSet
        If IntSN = "" Then
            Return Nothing
            Exit Function
        End If
        Using da As DataAccess = GetDataAccess()
            Dim StrSQL As String
            Try
                Dim dsTmp As New DataSet
                Dim dsFG As New DataSet
                Dim sTmp As String
                Dim daTmp As SqlDataAdapter
                Dim conSQL As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceAdditionConnString"))
                StrSQL = "exec sp_GetExtLaserSN '" + IntSN + "'"
                conSQL.Open()
                daTmp = New SqlDataAdapter(StrSQL, conSQL)
                daTmp.Fill(dsTmp, "tblWIP")
                If dsTmp Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables("tblWIP").Rows.Count < 1 Then
                    Return Nothing
                    Exit Function
                End If
                sTmp = dsTmp.Tables("tblWIP").Rows(0)(0).ToString.Trim + "&&" + dsTmp.Tables("tblWIP").Rows(0)(1).ToString.Trim
                StrSQL = String.Format("exec SP_GetWIPShipSN '{0}'", sTmp)
                dsFG = da.ExecuteDataSet(StrSQL)
                If dsFG Is Nothing Then
                    Return dsTmp
                    Exit Function
                End If
                If dsFG.Tables Is Nothing Then
                    Return dsTmp
                    Exit Function
                End If
                dsFG.Tables(0).TableName = "tblFG"
                dsTmp.Tables.Add(dsFG.Tables(0).Copy)
                Return dsTmp
            Catch ex As Exception
                ErrorLogging("SFC-GetWIPFGSN", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function SaveDJ(ByVal DJ As DataSet) As Boolean
        SaveDJ = False
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceAdditionConnString"))
        Dim sSql As String
        myConn.Open()
        Dim myTrans As SqlClient.SqlTransaction = myConn.BeginTransaction()
        Dim myTransCommand As SqlClient.SqlCommand = myConn.CreateCommand()
        Dim iRow As Integer
        myTransCommand.Transaction = myTrans
        Try
            Try
                sSql = String.Format("Delete from ST_DJHead where Createdon is not null")
                myTransCommand.CommandText = sSql
                myTransCommand.ExecuteNonQuery()
                For iRow = 0 To DJ.Tables(0).Rows.Count - 1
                    sSql = "insert into ST_DJHead(DJ,OrgCode,Assembly,Size,Status,ReleaseDate,ProductionLine) select "
                    sSql = sSql + "'" + DJ.Tables("DJList").Rows(iRow)("DJ#").ToString.Trim + "','" + DJ.Tables("DJList").Rows(iRow)("OrgCode").ToString.Trim + "',"
                    sSql = sSql + "'" + DJ.Tables("DJList").Rows(iRow)("Assembly").ToString.Trim + "'," + DJ.Tables("DJList").Rows(iRow)("DJSize").ToString.Trim + ","
                    sSql = sSql + "'" + DJ.Tables("DJList").Rows(iRow)("Status").ToString.Trim + "','" + DJ.Tables("DJList").Rows(iRow)("Release_Date").ToString.Trim + "',"
                    sSql = sSql + "'" + DJ.Tables("DJList").Rows(iRow)("Production_Line").ToString.Trim + "'"
                    myTransCommand.CommandText = sSql
                    myTransCommand.ExecuteNonQuery()
                Next
                myTrans.Commit()
                Return True
            Catch ex As Exception
                myTrans.Rollback()
            End Try
        Catch ex As Exception
            myTrans.Rollback()
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    Public Function ReadOrg(ByVal ServerName As String) As String
        ReadOrg = ""
        Dim cmdCon As SqlClient.SqlCommand
        Dim adoCon As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim sSQL, sOrg As String
        Try
            adoCon.Open()
            sSQL = "Select distinct ltrim(rtrim(OrgCode)) from T_Server where ServerName ='" + ServerName + "'"
            cmdCon = New SqlClient.SqlCommand(sSQL, adoCon)
            objReader = cmdCon.ExecuteReader()
            sOrg = ""
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then sOrg = sOrg + objReader.GetValue(0) + ","
            End While
            If sOrg.Substring(sOrg.Length - 1) = "," Then
                sOrg = sOrg.Substring(0, sOrg.Length - 1)
            End If
            ReadOrg = sOrg
        Catch ex As Exception
            ReadOrg = ""
        Finally
            If adoCon.State = ConnectionState.Open Then
                adoCon.Close()
            End If
        End Try
    End Function

    Public Function getMessageBMessageCode(ByVal MessageCode As String, ByVal Lang As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_getMessageByMessagecode '{0}','{1}'", MessageCode, Lang)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-getMessageByMessagecode", "", ex.Message & ex.Source, "E")
                Return "Get Message Fail."
            End Try
        End Using
    End Function

    Public Function GetModelByExtSN(ByVal ExtSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetModelByExtSN '{0}'", ExtSN)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-GetModelByExtSN", "", ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetWIPTestData(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dsTmp As New DataSet
            Dim dsReturn As New DataSet
            Try
                strSQL = String.Format("exec sp_GetWIPTestData '{0}','{1}'", IntSN, 1)
                dsTmp = Nothing
                dsTmp = da.ExecuteDataSet(strSQL)
                If dsTmp Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables(0).Rows.Count < 1 Then
                    Return Nothing
                    Exit Function
                End If
                dsTmp.Tables(0).TableName = "tblWIPHeader"
                dsReturn.Tables.Add(dsTmp.Tables(0).Copy)
                strSQL = String.Format("exec sp_GetWIPTestData '{0}','{1}'", IntSN, 2)
                dsTmp = Nothing
                dsTmp = da.ExecuteDataSet(strSQL)
                If dsTmp Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                dsTmp.Tables(0).TableName = "tblWIPFlow"
                dsReturn.Tables.Add(dsTmp.Tables(0).Copy)
                strSQL = String.Format("exec sp_GetWIPTestData '{0}','{1}'", IntSN, 3)
                dsTmp = Nothing
                dsTmp = da.ExecuteDataSet(strSQL)
                If dsTmp Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                dsTmp.Tables(0).TableName = "tblTDHeader"
                dsReturn.Tables.Add(dsTmp.Tables(0).Copy)
                strSQL = String.Format("exec sp_GetWIPTestData '{0}','{1}'", IntSN, 4)
                dsTmp = Nothing
                dsTmp = da.ExecuteDataSet(strSQL)
                If dsTmp Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                If dsTmp.Tables Is Nothing Then
                    Return Nothing
                    Exit Function
                End If
                dsTmp.Tables(0).TableName = "tblTDItem"
                dsReturn.Tables.Add(dsTmp.Tables(0).Copy)
                Return dsReturn
            Catch ex As Exception
                ErrorLogging("SFC-GetWIPTestData", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function IsReworkUnit(ByVal WIPID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, result As String
            Try
                strSQL = String.Format("exec sp_IsReworkUnit '{0}'", WIPID)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
                Return result
            Catch ex As Exception
                ErrorLogging("SFC-IsReworkUnit", "", ex.Message & ex.Source, "E")
                Return "N"
            End Try
        End Using
    End Function

    Public Function IsLastProcess(ByVal Model As String, ByVal PCBA As String, ByVal Process As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, result As String
            Try
                strSQL = String.Format("exec sp_IsLastProcess '{0}','{1}','{2}'", Model, PCBA, Process)
                result = Convert.ToString(da.ExecuteScalar(strSQL))
                Return result
            Catch ex As Exception
                ErrorLogging("SFC-IsLastProcess", "", ex.Message & ex.Source, "E")
                Return "N"
            End Try
        End Using
    End Function

    Public Function IsBottomPCBA(ByVal Model As String, ByVal PCBA As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, result As String
            Try
                strSQL = String.Format("exec sp_IsBottomPCBA '{0}','{1}'", Model, PCBA)
                result = da.ExecuteScalar(strSQL).ToString
                Return result
            Catch ex As Exception
                ErrorLogging("SFC-IsBottomPCBA", "", ex.Message & ex.Source, "E")
                Return "N"
            End Try
        End Using
    End Function


    Public Function GetMaintainExpireLog(ByVal Type As String, ByVal Expire As Integer, Optional ByVal FixtureID As String = "") As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_GetMaintainExpireLog '{0}','{1}','{2}'", FixtureID, Type, Expire)
            Return da.ExecuteDataSet(strSQL).GetXml
        Catch ex As Exception
            ErrorLogging("SFC-GetMaintainExpireLog", Type, ex.Message & ex.Source, "E")
            Return ""
        End Try
    End Function

    Public Function WIP_UpdateParameter(ByVal IntSN As String, ByVal Process As String, ByVal TestStep As String, ByVal TestName As String, ByVal LowerLimit As Double, ByVal UperLimit As Double) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_TempGapStepInquiry '{0}','{1}','{2}','{3}','{4}',{5}", IntSN, Process, TestStep, TestName, LowerLimit, UperLimit)
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-WIPUpdateParameter", "", ex.Message & ex.Source, "E")
            Return "3"
        End Try
    End Function

    Public Function Temp_InquiryProcess(ByVal IntSN As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_TempGapStepInquiry '{0}'", IntSN)
            Return da.ExecuteDataSet(strSQL)
        Catch ex As Exception
            ErrorLogging("SFC-TempInquiryProcess", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Temp_UpdateProcess(ByVal IntSN As String, ByVal SeqNo As Integer, ByVal Process As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_TempGapStep '{0}',{1},'{2}'", IntSN, SeqNo, Process)
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-TempUpdateProcess", "", ex.Message & ex.Source, "E")
            Return "0"
        End Try
    End Function

    Public Function GetMacAddress() As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_getMacAdress '{0}'", "TE")
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-GetMacAddress", "", ex.Message & ex.Source, "E")
            Return "0"
        End Try
    End Function

    Public Function SNListChangeBox(ByVal SNList As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ChangeBoxByBatch N'{0}'", DStoXML(SNList))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-SNListChangeBox", "", ex.Message)
            End Try
            Return result
        End Using
    End Function

    Public Function ChangeRevision(ByVal SNList As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ChangeRevision N'{0}'", DStoXML(SNList))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-ChangeRevision", "", ex.Message)
            End Try
            Return result
        End Using
    End Function


    Public Function getCartonInfo(ByVal boxid As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_getCartonInfo '{0}'", boxid)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("SFC-getCartonInfo", "", ex.Message)
                Return New DataSet
            End Try
        End Using
    End Function


    Public Function GetShipInfoByBoxIDSN(ByVal BoxIDSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet
            Try
                strSQL = String.Format("exec sp_getSNList_boxid_sn '{0}'", BoxIDSN)
                DS = da.ExecuteDataSet(strSQL)
                Return DS
            Catch ex As Exception
                ErrorLogging("SFC-GetShipInfoByBoxIDSN", "", ex.Message)
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function CheckCompIssueToDJ(ByVal DJ As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_Matching1_CheckCompIssueToDJ '{0}'", DJ)
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-CheckCompIssueToDJ", "", ex.Message)
            Return "0"
        End Try

    End Function

    Public Function IntSNPattern(ByVal Model As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_Matching1_intSNPattern '{0}'", Model)
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-IntSNPattern", "", ex.Message)
            Return ""
        End Try

    End Function

    Public Function GetResultAndPCBAList(ByVal IntSN As String, ByVal Proc As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_GetResultAndPCBAList '{0}','{1}'", IntSN, Proc)
                Return da.ExecuteDataSet(strSQL, 1)
            Catch ex As Exception
                ErrorLogging("SFC-GetResultAndAttributesList", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function WIPMatchingN(ByVal DSWIP As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_WIPMatchingN N'{0}'", DStoXML(DSWIP))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-WIPMatchingN", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function CheckFixtureID(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_CheckFixtureID N'{0}'", FixtureID)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("sp_CheckFixtureID", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function
    Public Function GetPanelIDByIntSN(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_GetPanelIDByIntSN N'{0}'", IntSN)
                'ErrorLogging("GetPanelIDByIntSN", "", "strSQL:" + strSQL, "D")
                result = da.ExecuteScalar(strSQL).ToString
                'ErrorLogging("GetPanelIDByIntSN", "", "result:" + result, "D")
            Catch ex As Exception
                ErrorLogging("sp_GetPanelIDByIntSN", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function
    Public Function GetWipHeaderByIntSN(ByVal intSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                Dim ds As New DataSet()
                sqlstr = String.Format("exec sp_GetWipHeadrByIntSN '{0}'", intSN)
                ds = da.ExecuteDataSet(sqlstr, "WipInfo")
                Return ds
            Catch ex As Exception
                Throw ex
                ErrorLogging("TDCService-GetOrderInfoFromETRACE", "", "intSN: " & intSN & ", " & ex.Message & ex.Source)
            End Try
        End Using

    End Function
    Public Function ClearFixtureID(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_ClearFixtureID  N'{0}'", FixtureID)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-ClearFixtureID", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function
    Public Function FixtureMount(ByVal DsFixture As DataSet, ByVal fixtureId As String, ByVal user As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = String.Empty
            Try
                strSQL = String.Format("exec sp_FixTureMount N'{0}',{1},{2}", DStoXML(DsFixture), fixtureId, user)
                ErrorLogging("SFC-FixTureMount", "", strSQL, "E")
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("SFC-FixTureMount", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function TE_IntSNListReadbyFixtureID(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_TE_IntSNListReadbyFixtureID N'{0}'", FixtureID)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("TE_IntSNListReadbyFixtureID", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function TE_ReworkUnitFlag(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_TE_ReworkUnitFlag N'{0}'", IntSN)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("TE_ReworkUnitFlag", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function TE_UnbindFixtureID(ByVal FixtureID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_TE_UnbindFixtureID N'{0}'", FixtureID)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("TE_UnbindFixtureID", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

#End Region

#Region "Smart Card Validation"
    Public Function SaveSmartCardHistory(ByVal ParamArray cardParams As String()) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                strSQL = String.Format("exec sp_SaveSmartCardHistory N'{0}', '{1}', '{2}','{3}','{4}','{5}',N'{6}', '{7}'", _
                cardParams(0), cardParams(1), cardParams(2), cardParams(3), cardParams(4), cardParams(5), cardParams(6), cardParams(7))
                Dim dt As DataTable = da.ExecuteDataSet(strSQL, "SmartCardInfo").Tables(0)
                Return dt
            Catch ex As Exception
                ErrorLogging("TDC-WS-SaveSmartCardHistory", "", ex.Message & ex.Source, "E")
                Return New DataTable()
            End Try
            Return New DataTable()
        End Using
    End Function


#End Region


#Region "ZS"

#End Region

#Region "FY"
    Public Function Touch_SaveRawData(ByVal PCBSN As String, ByVal Block As Integer, ByVal CircuitCode As String, ByVal SymptomCode As String, ByVal Model As String, ByVal PCBType As String, ByVal Stage As String, ByVal Result As String, ByVal Shift As String, ByVal Line As String, ByVal OperatorName As String, ByVal ATENo As String, ByVal WSEquipment As String) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Dim sSql As String
        Try
            Dim sReturn As String
            sSql = "exec sp_TouchSaveRawData '" + PCBSN + "','" + Block.ToString + "','" + CircuitCode + "',N'" + SymptomCode + "','" + Model + "','" + PCBType + "','" + Stage + "','" + Result + "','" + Shift + "','" + Line + "',N'" + OperatorName + "','" + ATENo + "','" + WSEquipment + "'"
            myConn.Open()
            Dim myCommand As SqlClient.SqlCommand = New SqlCommand(sSql, myConn)
            sReturn = myCommand.ExecuteScalar.ToString()
            sSql = "exec sp_TouchReport '" + CircuitCode + "',N'" + SymptomCode + "','" + Model + "','" + PCBType + "','" + Stage + "','" + Result + "','" + Shift + "','" + Line + "',N'" + OperatorName + "','" + ATENo + "','" + WSEquipment + "'"
            Dim QAConn As SqlConnection = New SqlConnection(ConfigurationManager.AppSettings("eTraceQAConnString"))
            QAConn.Open()
            Dim QAmyCommand As SqlClient.SqlCommand = New SqlCommand(sSql, QAConn)
            QAmyCommand.ExecuteScalar.ToString()
            If sReturn = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ErrorLogging("SFC_TouchSaveRawData", "", ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function Touch_ReadRawData(ByVal StartDate As String, ByVal EndDate As String) As String
        Dim sSql As String
        sSql = "select PCBSN,Block,CircuitCode,SymptomCode,ProductionTime as ProductTime,Model,PCBType,Stage,Result,Shift,Line,Operator,ATENo,WSEquipment from T_TUPRaw where replace(convert(char(10),ProductionTime,121),'-','') between '" + StartDate + "' and '" + EndDate + "'"
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Try
            Dim dsTmp As New DataSet
            Dim daSQL As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(sSql, myConn)
            daSQL.Fill(dsTmp, "RawData")
            Return dsTmp.GetXml
        Catch ex As Exception
            ErrorLogging("SFC_TouchReadRawData", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Touch_SaveStaticsData(ByVal ProductTime As String, ByVal Line As String, ByVal CircuitCode As String, ByVal SymptomCode As String, ByVal PASSQty As Integer, ByVal FAILQty As Integer, ByVal Block As Integer) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Dim sSql As String
        Try
            Dim sReturn As String
            sSql = "exec sp_TouchSaveStaticsData '" + ProductTime + "','" + Line + "','" + CircuitCode + "',N'" + SymptomCode + "'," + PASSQty.ToString + "," + FAILQty.ToString + "," + Block.ToString + ""
            myConn.Open()
            Dim myCommand As SqlClient.SqlCommand = New SqlCommand(sSql, myConn)
            sReturn = myCommand.ExecuteScalar.ToString()
            If sReturn = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ErrorLogging("SFC_TouchSaveStaticsData", "", ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function Touch_ReadStaticsData(ByVal ProdTime As String, ByVal LineID As String) As String
        Dim sSql As String
        sSql = "select ProductionDate as ProductTime,Line,CircuitCode,SymptomCode,PASSQty,FAILQty,Block from T_TUPStatics where ProductionDate='" + ProdTime + "' and Line='" + LineID + "'"
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Try
            Dim dsTmp As New DataSet
            Dim daSQL As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(sSql, myConn)
            daSQL.Fill(dsTmp, "StaticsData")
            Return dsTmp.GetXml
        Catch ex As Exception
            ErrorLogging("SFC_TouchReadStaticsData", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Touch_SaveCAData(ByVal ProductTime As String, ByVal Line As String, ByVal Description As String, ByVal CircuitCode As String, ByVal SymptomCode As String, ByVal Reason As String, ByVal ImproveAction As String, ByVal ActionOwner As String, ByVal DueDate As DateTime, ByVal Status As String) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Dim sSql As String
        Try
            Dim sReturn As String
            sSql = "exec sp_TouchSaveCAData '" + ProductTime + "','" + Line + "',N'" + Description + "','" + CircuitCode + "',N'" + SymptomCode + "',N'" + Reason + "',N'" + ImproveAction + "',N'" + ActionOwner + "','" + DueDate.ToString + "','" + Status + "'"
            myConn.Open()
            Dim myCommand As SqlClient.SqlCommand = New SqlCommand(sSql, myConn)
            sReturn = myCommand.ExecuteScalar.ToString()
            If sReturn = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ErrorLogging("SFC_TouchSaveCAData", "", ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function Touch_UpdateCAData(ByVal CANo As String, ByVal Reason As String, ByVal ImproveAction As String, ByVal ActionOwner As String, ByVal DueDate As DateTime, ByVal Status As String) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Dim sSql As String
        Try
            Dim sReturn As String
            sSql = "exec sp_TouchUpdateCAData '" + CANo + "',N'" + Reason + "',N'" + ImproveAction + "',N'" + ActionOwner + "','" + DueDate + "','" + Status + "'"
            myConn.Open()
            Dim myCommand As SqlClient.SqlCommand = New SqlCommand(sSql, myConn)
            sReturn = myCommand.ExecuteScalar.ToString()
            If sReturn = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ErrorLogging("SFC_TouchUpdateCAData", "", ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function Touch_ReadCAData(ByVal ProductTime As String, ByVal LineID As String, ByVal Status As String) As String
        Dim sSql As String
        sSql = "select CANo,ProductionDate as ProductTime,Line,Description,CircuitCode,SymptomCode,Reason,ImproveAction,ActionOwner,Status from T_TUPCA where ProductionDate='" + ProductTime + "' and Line='" + LineID + "' and Status='" + Status + "'"
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceOtherConnString"))
        Try
            Dim dsTmp As New DataSet
            Dim daSQL As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(sSql, myConn)
            daSQL.Fill(dsTmp, "CAData")
            Return dsTmp.GetXml
        Catch ex As Exception
            ErrorLogging("SFC_TouchReadCAData", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function
#End Region

#Region "PH"

#End Region

#Region "Flatfile"

    Public Function getFlatFileProperties(ByVal type As String, ByVal attribute As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim value As String
            Try
                strsql = String.Format("select " + attribute + " from T_FlatFileProperties WITH(NOLOCK) where propertiesType ='" + type + "'")
                value = da.ExecuteScalar(strsql)
            Catch ex As Exception
                ErrorLogging("TDC-getFlatFileProperties", "", ex.Message & ex.Source, "E")
                value = "No record"
            End Try
            Return value
        End Using
    End Function


    Public Function setFlatFileProperties(ByVal type As String, ByVal attribute As String, ByVal value As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Try
                strsql = String.Format("update  T_FlatFileProperties set " + attribute + "= '" + value + "' where  propertiesType ='" + type + "'")
                da.ExecuteNonQuery(strsql)
                Return True
            Catch ex As Exception
                ErrorLogging("TDC-setFlatFileProperties", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function


    Public Function SaveFlatFileMessage(ByVal model As String, ByVal boxid As String, ByVal palletid As String, ByVal serialno As String, ByVal flatfile As String, ByVal sentby As String) As Boolean

        Using da As DataAccess = GetDataAccess()
            Dim Ds As New Data.DataSet
            Dim sqlStr As String

            sqlStr = "Update T_Shippment set FlatFile='" + flatfile + "',Sentby='" + sentby + "' where model='" & model & "'"
            If boxid <> "" Then
                sqlStr = sqlStr + " and CartonID in('" & boxid.Replace(",", "','") & "')"
            End If
            If palletid <> "" Then
                sqlStr = sqlStr + " and palletid in ('" & palletid.Replace(",", "','") & "')"
            End If
            If serialno <> "" Then
                sqlStr = sqlStr + " and Productserialno in ('" & serialno.Replace(",", "','") & "')"
            End If

            Try
                da.ExecuteNonQuery(String.Format(sqlStr))
                Return True
            Catch ex As Exception
                ErrorLogging("TDC-SaveFlatFileMessage", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function saveFlatFileSN(ByVal FlatfileDS As DataSet, ByVal username As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim strSQL As String

            Dim i As Integer
            Try
                For i = 0 To FlatfileDS.Tables("FlatFile").Rows.Count - 1
                    strSQL = String.Format("insert T_FlatfileSN(ProductSerialNo,SerialNo2,CreatedBy)values ('{0}','{1}','{2}')", FlatfileDS.Tables("FlatFile").Rows(i).Item(0), FlatfileDS.Tables("FlatFile").Rows(i).Item(5), username)
                    da.ExecuteNonQuery(strSQL)
                Next
                Return True
            Catch ex As Exception
                ErrorLogging("FlatFile-saveFlatFileSN", username, ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function


    Public Function saveFlatFileSN(ByVal FlatfileDS As DataSet, ByVal ParamArray parmsArray As String()) As Boolean
        ''Dim cmdSQL As SqlClient.SqlCommand
        ''Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String = String.Empty, msg As String = String.Empty
        ''Dim change As String
        ''Dim i As Integer
        Using da As DataAccess = GetDataAccess()

            Try
                ''myConn.Open()

                strSQL = String.Format("exec sp_SaveFlatFileSN '{0}', '{1}', '{2}', N'{3}'", DStoXML(FlatfileDS), parmsArray(0), parmsArray(1), parmsArray(2))
                ''cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                ''msg = Convert.ToString(cmdSQL.ExecuteScalar)
                ''ErrorLogging("TDC-saveFlatFileSN", parmsArray(2), msg)
                ''myConn.Close()
                msg = da.ExecuteScalar(strSQL)
                If (Not String.IsNullOrEmpty(msg)) Then
                    Return False
                End If
                Return True
            Catch ex As Exception
                ErrorLogging("TDC-saveFlatFileSN", parmsArray(2), ex.Message)
                Return False
                ''Finally
                ''    If myConn.State = ConnectionState.Open Then
                ''        myConn.Close()
                ''    End If
            End Try

        End Using
    End Function

    ''' <summary>
    ''' Update by Jackson Huang
    ''' </summary>
    ''' <param name="model"></param>
    ''' <param name="BoxID"></param>
    ''' <param name="PalletID"></param>
    ''' <param name="SerialNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''Public Function GetFlatfile(ByVal model As String, ByVal BoxID As String, ByVal PalletID As String, ByVal SerialNo As String) As Data.DataSet
    ''    Using da As DataAccess = GetDataAccess()
    ''        Dim Ds As New Data.DataSet
    ''        Dim sqlStr As String
    ''        Try
    ''            sqlStr = "SELECT ProductSerialNo, Model, CustomerPN, CustomerRev,CreatedOn,SerialNo2, SerialNo3, SerialNo4, FlatFile FROM T_Shippment WITH (NOLOCK) where model='" & model & "'"
    ''            If BoxID <> "" Then
    ''                sqlStr = sqlStr + " and CartonID in('" & BoxID.Replace(",", "','") & "')"
    ''            End If
    ''            If PalletID <> "" Then
    ''                sqlStr = sqlStr + " and palletid in ('" & PalletID.Replace(",", "','") & "')"
    ''            End If
    ''            If SerialNo <> "" Then
    ''                sqlStr = sqlStr + " and Productserialno in ('" & SerialNo.Replace(",", "','") & "')"
    ''            End If
    ''            sqlStr = sqlStr + " order by createdon desc"
    ''            Ds = da.ExecuteDataSet(String.Format(sqlStr), "FlatFile")
    ''            Return Ds
    ''        Catch ex As Exception
    ''            ErrorLogging("FlatFile-GetFlatfile", sqlStr, ex.Message)
    ''        End Try
    ''    End Using
    ''End Function



    Public Function GetFlatfile(ByVal model As String, ByVal BoxID As String, ByVal PalletID As String, ByVal SerialNo As String) As Data.DataSet
        'Dim cmdSQL As SqlClient.SqlCommand
        ''Dim Conn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        'Dim strSQL As String
        'Dim objReader As SqlClient.SqlDataReader
        ' Dim PalletID As String = Nothing
        ''Dim Rs As New Data.SqlClient.SqlDataAdapter
        Dim Ds As New Data.DataSet
        Dim errorM As String
        Dim sqlStr As String
        ''Or model = "AA24690RS5"
        If (model = "700-004750-0000" Or model = "700-004751-0000" Or model = "AA27280L" Or model = "AA23540RS5") Then
            sqlStr = "SELECT ProductSerialNo, Model, CustomerPN, CustomerRev,CreatedOn,'' as SerialNo2, SerialNo3, SerialNo4, FlatFile FROM T_Shippment with (nolock) where model='" & model & "'"
        Else
            sqlStr = "SELECT ProductSerialNo, Model, CustomerPN, CustomerRev,CreatedOn,SerialNo2, SerialNo3, SerialNo4, FlatFile FROM T_Shippment with (nolock) where model='" & model & "'"
        End If
        'sqlStr = "SELECT ProductSerialNo, Model, CustomerPN, CustomerRev,CreatedOn,SerialNo2,FlatFile FROM T_Shippment where model='" & model & "'"
        If BoxID <> "" Then
            sqlStr = sqlStr + " and CartonID in('" & BoxID.Replace(",", "','") & "')"
        End If
        If PalletID <> "" Then
            sqlStr = sqlStr + " and palletid in ('" & PalletID.Replace(",", "','") & "')"
        End If
        If SerialNo <> "" Then
            sqlStr = sqlStr + " and Productserialno in ('" & SerialNo.Replace(",", "','") & "')"
        End If
        sqlStr = sqlStr + " order by createdon desc"
        Using da As DataAccess = GetDataAccess()

            Try
                Ds = da.ExecuteDataSet(sqlStr)
                Ds.Tables(0).TableName = "FlatFile"
                ''Rs.Fill(Ds, "FlatFile")
                ''Using da As DataAccess = GetDataAccess()
                Dim strSQL As String = String.Format("exec sp_GetTransParameter '{0}', '{1}'", model, String.Empty)
                Dim dt As DataTable = da.ExecuteDataTable(strSQL)
                If (dt.Rows.Count > 0) Then
                    dt.TableName = "TransParaInfo"
                    Ds.Tables.Add(dt)
                End If

                ''End Using


                Return Ds
            Catch ex As Exception
                ErrorLogging("TDC-GetFlatfile", "", ex.Message + " " + sqlStr)
                Return Ds
            End Try
        End Using
    End Function

    Public Function CheckFlatFileSN(ByVal FlatfileDS As DataSet, ByVal username As String) As String
        ''Dim cmdSQL As SqlClient.SqlCommand
        ''Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String
        CheckFlatFileSN = ""
        Dim i As Integer
        Using da As DataAccess = GetDataAccess()


            Try
                ''myConn.Open()

                strSQL = String.Format("exec sp_CheckFlatFileSN '{0}'", DStoXML(FlatfileDS))
                ''ErrorLogging("TDC-CheckFlatFileSN", username, DStoXML(FlatfileDS))
                ''cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                CheckFlatFileSN = da.ExecuteScalar(strSQL)
                ''ErrorLogging("TDC-CheckFlatFileSN", username, CheckFlatFileSN)
                ''myConn.Close()

            Catch ex As Exception
                ErrorLogging("TDC-CheckFlatFileSN", username, ex.Message)

                ''Finally
                ''    If myConn.State = ConnectionState.Open Then
                ''        myConn.Close()
                ''    End If
            End Try
        End Using
    End Function

#End Region

#Region "QCcode"

    Public Function SaveTDCRepairCode(ByVal RepairCode As DataSet) As Boolean
        Dim InsertCommand As SqlClient.SqlCommand
        Dim UpdateCode As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As String
        InsertCommand = New SqlClient.SqlCommand("Insert into T_RepCodes (Site,Category,CodeGroup,Code,Description,Remarks,Status,CompReplacement,ReturnTo) values (@Site,@Category,@CodeGroup,@Code,@Description,@Remarks,@Status,@CompReplacement,@ReturnTo) ", myConn)
        InsertCommand.Parameters.Add("@Site", SqlDbType.VarChar, 50, "Site")
        InsertCommand.Parameters.Add("@Category", SqlDbType.VarChar, 50, "Category")
        InsertCommand.Parameters.Add("@CodeGroup", SqlDbType.VarChar, 50, "CodeGroup")
        InsertCommand.Parameters.Add("@Code", SqlDbType.VarChar, 50, "Code")
        InsertCommand.Parameters.Add("@Description", SqlDbType.NVarChar, 100, "Description")
        InsertCommand.Parameters.Add("@Remarks", SqlDbType.NVarChar, 100, "Remarks")
        InsertCommand.Parameters.Add("@Status", SqlDbType.Int, 50, "Status")
        InsertCommand.Parameters.Add("@CompReplacement", SqlDbType.Bit, 1, "CompReplacement")
        InsertCommand.Parameters.Add("@ReturnTo", SqlDbType.VarChar, 100, "ReturnTo")

        UpdateCode = New SqlClient.SqlCommand("Update T_RepCodes SET Description=@Description, Remarks=@Remarks, Status=@Status, CompReplacement=@CompReplacement, ReturnTo=@ReturnTo where Code=@Code and CodeGroup=@CodeGroup", myConn)
        UpdateCode.Parameters.Add("@Code", SqlDbType.VarChar, 50, "Code")
        UpdateCode.Parameters.Add("@CodeGroup", SqlDbType.VarChar, 50, "CodeGroup")
        UpdateCode.Parameters.Add("@Description", SqlDbType.NVarChar, 100, "Description")
        UpdateCode.Parameters.Add("@Remarks", SqlDbType.NVarChar, 100, "Remarks")
        UpdateCode.Parameters.Add("@Status", SqlDbType.Int, 50, "Status")
        UpdateCode.Parameters.Add("@CompReplacement", SqlDbType.Bit, 1, "CompReplacement")
        UpdateCode.Parameters.Add("@ReturnTo", SqlDbType.VarChar, 100, "ReturnTo")
        Try
            myConn.Open()
            Dim i As Integer
            For i = 0 To RepairCode.Tables(0).Rows.Count - 1
                If Not RepairCode.Tables(0).Rows(i)("Code") Is DBNull.Value Then
                    If Not RepairCode.Tables(0).Rows(i)("Status") Is DBNull.Value Then
                        'Update
                        If Convert.ToString(RepairCode.Tables(0).Rows(i)("Description")) <> Convert.ToString(RepairCode.Tables(1).Rows(i)("Description")) Or Convert.ToString(RepairCode.Tables(0).Rows(i)("Remarks")) <> Convert.ToString(RepairCode.Tables(1).Rows(i)("Remarks")) Or RepairCode.Tables(0).Rows(i)("Status") <> RepairCode.Tables(1).Rows(i)("Status") _
                            Or RepairCode.Tables(0).Rows(i)("CompReplacement") <> RepairCode.Tables(1).Rows(i)("CompReplacement") Or RepairCode.Tables(0).Rows(i)("ReturnTo").ToString() <> RepairCode.Tables(1).Rows(i)("ReturnTo").ToString() Then

                            UpdateCode.Parameters("@Code").Value = RepairCode.Tables(0).Rows(i)("Code")
                            UpdateCode.Parameters("@CodeGroup").Value = RepairCode.Tables(0).Rows(i)("CodeGroup")
                            UpdateCode.Parameters("@Description").Value = RepairCode.Tables(0).Rows(i)("Description")
                            UpdateCode.Parameters("@Remarks").Value = RepairCode.Tables(0).Rows(i)("Remarks")
                            UpdateCode.Parameters("@Status").Value = RepairCode.Tables(0).Rows(i)("Status")
                            UpdateCode.Parameters("@CompReplacement").Value = RepairCode.Tables(0).Rows(i)("CompReplacement")
                            UpdateCode.Parameters("@ReturnTo").Value = RepairCode.Tables(0).Rows(i)("ReturnTo").ToString()
                            UpdateCode.CommandType = CommandType.Text
                            ra = UpdateCode.ExecuteNonQuery()
                        End If
                    Else
                        'Insert
                        InsertCommand.Parameters("@Site").Value = RepairCode.Tables(0).Rows(i)("Site")
                        InsertCommand.Parameters("@Category").Value = RepairCode.Tables(0).Rows(i)("Category")
                        InsertCommand.Parameters("@CodeGroup").Value = RepairCode.Tables(0).Rows(i)("CodeGroup")
                        InsertCommand.Parameters("@Code").Value = RepairCode.Tables(0).Rows(i)("Code")
                        InsertCommand.Parameters("@Description").Value = RepairCode.Tables(0).Rows(i)("Description")
                        InsertCommand.Parameters("@Remarks").Value = RepairCode.Tables(0).Rows(i)("Remarks")
                        InsertCommand.Parameters("@Status").Value = "1"
                        InsertCommand.Parameters("@CompReplacement").Value = RepairCode.Tables(0).Rows(i)("CompReplacement")
                        InsertCommand.Parameters("@ReturnTo").Value = RepairCode.Tables(0).Rows(i)("ReturnTo").ToString()
                        InsertCommand.CommandType = CommandType.Text
                        ra = InsertCommand.ExecuteNonQuery()
                    End If
                End If
            Next
            SaveTDCRepairCode = True
        Catch ex As Exception
            ErrorLogging("TDC-SaveTDCRepairCode", "", ex.Message & ex.Source, "E")
            SaveTDCRepairCode = False
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function


    Public Function ReadRepairCodeForQCcode(ByVal Category As String, ByVal Site As String) As DataSet

        ' First, setup the dataset
        ReadRepairCodeForQCcode = New DataSet
        Dim RepairCode As DataTable = New Data.DataTable("RepairCode")
        Dim RepairDataRow As Data.DataRow
        Dim myDataColumn As DataColumn

        myDataColumn = New Data.DataColumn("Site", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Category", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CodeGroup", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Code", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Description", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Remarks", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.Int16"))
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CompReplacement", System.Type.GetType("System.Boolean"))
        myDataColumn.DefaultValue = False
        RepairCode.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReturnTo", System.Type.GetType("System.String"))
        RepairCode.Columns.Add(myDataColumn)
        ReadRepairCodeForQCcode.Tables.Add(RepairCode)

        Dim OriginalDataCode As DataTable = New Data.DataTable("OriginalDataCode")
        Dim OriginalDataRow As Data.DataRow
        Dim OriginalDataColumn As DataColumn

        OriginalDataColumn = New Data.DataColumn("Site", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("Category", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("CodeGroup", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("Code", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("Description", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("Remarks", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.Int16"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("CompReplacement", System.Type.GetType("System.Boolean"))
        OriginalDataColumn.DefaultValue = False
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        OriginalDataColumn = New Data.DataColumn("ReturnTo", System.Type.GetType("System.String"))
        OriginalDataCode.Columns.Add(OriginalDataColumn)
        ReadRepairCodeForQCcode.Tables.Add(OriginalDataCode)

        Dim ErrorTable As DataTable
        Dim ErrorRow As Data.DataRow
        ErrorTable = New Data.DataTable("ErrorTable")
        myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
        ErrorTable.Columns.Add(myDataColumn)
        ReadRepairCodeForQCcode.Tables.Add(ErrorTable)

        'Now read data from database
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        CLMasterSQLCommand = New SqlClient.SqlCommand("Select T_RepCodes.Site ,T_RepCodes.Category ,T_RepCodes.CodeGroup ,T_RepCodes.Code ,T_RepCodes.Description ,T_RepCodes.Remarks ,T_RepCodes.Status,CompReplacement,ReturnTo from T_RepCodes  Where T_RepCodes.Site like @Site and T_RepCodes.Category like @Category ", myConn)
        CLMasterSQLCommand.Parameters.Add("@Site", SqlDbType.VarChar, 50, "Site")
        CLMasterSQLCommand.Parameters("@Site").Value = Site
        CLMasterSQLCommand.Parameters.Add("@Category", SqlDbType.VarChar, 50, "Category")
        CLMasterSQLCommand.Parameters("@Category").Value = Category

        Try
            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                RepairDataRow = RepairCode.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then RepairDataRow("Site") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then RepairDataRow("Category") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then RepairDataRow("CodeGroup") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then RepairDataRow("Code") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then RepairDataRow("Description") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then RepairDataRow("Remarks") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then RepairDataRow("Status") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then RepairDataRow("CompReplacement") = objReader.GetValue(7)
                If Not objReader.GetValue(7) Is DBNull.Value Then RepairDataRow("ReturnTo") = objReader.GetValue(8)
                RepairCode.Rows.Add(RepairDataRow)
            End While
            myConn.Close()

            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                OriginalDataRow = OriginalDataCode.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then OriginalDataRow("Site") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then OriginalDataRow("Category") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then OriginalDataRow("CodeGroup") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then OriginalDataRow("Code") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then OriginalDataRow("Description") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then OriginalDataRow("Remarks") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then OriginalDataRow("Status") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then OriginalDataRow("CompReplacement") = objReader.GetValue(7)
                If Not objReader.GetValue(7) Is DBNull.Value Then OriginalDataRow("ReturnTo") = objReader.GetValue(8)
                OriginalDataCode.Rows.Add(OriginalDataRow)
            End While

            If RepairCode.Rows.Count = 0 Or OriginalDataCode.Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No record found!"
                ErrorTable.Rows.Add(ErrorRow)
            End If
        Catch ex As Exception
            ErrorLogging("TDC-ReadRepairCode", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function

    Public Function IsProductProcess(ByVal Process As String) As Boolean
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim result As Boolean = False
        Try
            strSQL = String.Format("exec sp_IsProductProcess '{0}'", Process)
            result = IIf(da.ExecuteScalar(strSQL).ToString = "Y", True, False)
        Catch ex As Exception
            ErrorLogging("SFC-ProductProcess", "", ex.Message)
        End Try
        Return result

    End Function
#End Region

    Public Function GetModelRevForTE(ByVal IntSN As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim result As String = ""
        Try
            strSQL = String.Format("exec sp_GetModelRevForTE '{0}'", IntSN)
            result = da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-GetModelRevForTE", "", ex.Message)
        End Try
        Return result

    End Function

    Public Function GetProdLineAndResultForTE(ByVal IntSN As String, ByVal Process As String) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim result As String = ""
        Try
            strSQL = String.Format("exec sp_GetProdLineAndResultForTE '{0}','{1}'", IntSN, Process)
            result = da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("SFC-GetProdLineAndResultForTE", "", ex.Message)
        End Try
        Return result

    End Function
    Public Function UploadProductRoutingLog(ByVal productRoutingLog As DataTable, ByVal userid As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String

            msg = String.Empty

            Try

                Dim dsRoutingLog As DataSet = New DataSet("RoutingLog")
                dsRoutingLog.Tables.Add(productRoutingLog)
                strSQL = String.Format("exec sp_InsertProductRoutingLog '{0}', '{1}'", DStoXML(dsRoutingLog), userid)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))

            Catch ex As Exception
                ErrorLogging("SFC UploadProductRouting", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function
    ''' <summary>
    ''' Upload Product change log for ProductMaster
    ''' </summary>
    ''' <param name="productRoutingLog"></param>
    ''' <param name="model"></param>
    ''' <returns></returns>
    Public Function UploadProductChangeLog(ByVal productRoutingLog As DataTable, ByVal model As String) As String

        Using da As DataAccess = GetDataAccess()
            Dim strSQL, msg As String
            msg = String.Empty
            Try
                Dim dsLog As DataSet = New DataSet("ProductChangeLog")
                dsLog.Tables.Add(productRoutingLog)
                strSQL = String.Format("exec sp_InsertProductChangeLog '{0}', '{1}'", DStoXML(dsLog), model)
                msg = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("SFC UploadProductChangeLog", "", ex.Message & ex.Source, "E")
                Return msg
            End Try
            Return msg
        End Using
    End Function
End Class
