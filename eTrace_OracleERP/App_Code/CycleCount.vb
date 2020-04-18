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



Public Structure GetCycleCount
    Dim Flag As String
    Dim Message As String
    Dim ds As DataSet
End Structure


Public Class CycleCount
    Inherits PublicFunction
#Region "SMTCycleCount"
    Public Function SMTCycleCountHH(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            Dim CLIDStatus As String = ""
            Dim vf As String
            Dim resultflag As String = ""
            Dim LightLed As Boolean
            Dim dsLed As New DataSet
            Dim dtSlot As New DataTable("dtSlot")
            dtSlot.Columns.Add(New Data.DataColumn("Slot", System.Type.GetType("System.String")))
            dsLed.Tables.Add(dtSlot)
            Dim SlotRow As Data.DataRow
            Dim myWMS As WMS = New WMS
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_SMTCycleCount_CLIDValid"
                sda.SelectCommand.Parameters.Add("@EventID", SqlDbType.VarChar, 150).Value = EventID
                sda.SelectCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 150).Value = CLID
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                If ds.Tables(0).Rows(0)("ErrMsgID") > 0 Then
                    Return ds
                Else
                    vf = SMTCLIDStatusChanged(EventID, CLID, CLIDStatus, resultflag)
                    dsLed.DataSetName = "DS"
                    SlotRow = dtSlot.NewRow()
                    SlotRow("Slot") = ds.Tables(0).Rows(0)("Slot")
                    dtSlot.Rows.Add(SlotRow)
                    Try
                        LightLed = myWMS.LEDControlBySlot(dsLed, 0, 0)
                    Catch ex As Exception
                        ErrorLogging("SMTCycleCountHH", OracleLoginData.User, ex.Message & ex.Source, "E")
                    End Try
                End If
            Catch ex As Exception
                ErrorLogging("SMTCycleCountHH", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
            Return ds
        End Using
    End Function
    Public Function SMTCycleCountCLIDValid(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_SMTCycleCount_CLIDValid"
                sda.SelectCommand.Parameters.Add("@EventID", SqlDbType.VarChar, 150).Value = EventID
                sda.SelectCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 150).Value = CLID
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("SMTCycleCountCLIDValid", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
            Return ds
        End Using
    End Function
    Public Function GetSMTCCName(ByVal OrgCode As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_SMTCycleCountName"
                sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 150).Value = OrgCode
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("SMTCycleCount-GetSMTCCName", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
            Return ds
        End Using
    End Function
    Public Function SMTCCSave(ByVal EventID As String, ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim sda As SqlClient.SqlDataAdapter
            Dim i As Integer
            Try
                For i = 0 To p_ds.Tables(0).Rows.Count - 1
                    If p_ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        p_ds.Tables(0).Rows(i).SetAdded()
                    End If
                Next
                sda = da.Sda_Insert()
                sda.InsertCommand.CommandType = CommandType.StoredProcedure
                sda.InsertCommand.CommandText = "sp_SMTCycleCount_Save"
                sda.InsertCommand.Parameters.Add("@EventID", SqlDbType.VarChar, 150).Value = EventID
                sda.InsertCommand.Parameters.Add("@SeqID", SqlDbType.Int).SourceColumn = "p_sequence"
                sda.InsertCommand.CommandTimeout = TimeOut_M5
                sda.InsertCommand.Connection.Open()

                sda.Update(p_ds.Tables(0))
                sda.InsertCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("SMTCycleCount-SMTCCSave", OracleLoginData.User, ex.Message & ex.Source, "E")
                SMTCCSave = "N"
            End Try
            SMTCCSave = "Y"
        End Using
    End Function
    Public Function GetSMTCCList(ByVal OrgCode As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_get_smtcyclecount"
                sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 150).Value = OrgCode
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("SMTCycleCount-GetSMTCCList", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
            Return ds
        End Using
    End Function
    Public Function GetSMTScanedCLID(ByVal EventID As String, ByVal Seq As Integer, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim Sqlstr As String
            Dim ds As New DataSet
            Try
                Sqlstr = String.Format("select a.CLID,a.Item,b.QtyBaseUOM as Qty,b.RTLot as Lot,a.InvSlot,a.Status from T_InvEventItem a inner join T_CLMaster b on a.CLID = b.CLID  where a.EventID = '{0}' and a.MachineID = {1}  ", EventID, Seq)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"CLID", "Item", "Qty", "Lot", "InvSlot", "Status"}
                ds = da.ExecuteDataSet(sql, tables)
            Catch ex As Exception
                ErrorLogging("SMTCycleCount-GetSMTScanedCLID", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
            Return ds
        End Using
    End Function
    Public Function SMTCLIDStatusChanged(ByRef EventID As String, ByRef CLID As String, ByRef Action As String, ByRef ResultFlag As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_SMTCycleCount_ChangedStatus"
                sda.SelectCommand.Parameters.Add("@EventID", SqlDbType.VarChar, 150).Value = EventID
                sda.SelectCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 150).Value = CLID
                sda.SelectCommand.Parameters.Add("@Action", SqlDbType.VarChar, 150).Value = Action
                sda.SelectCommand.Parameters.Add("@Result", SqlDbType.VarChar, 150)
                sda.SelectCommand.Parameters("@Result").Direction = ParameterDirection.Output
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                ResultFlag = sda.SelectCommand.Parameters("@Result").Value
                SMTCLIDStatusChanged = ResultFlag
            Catch ex As Exception
                ErrorLogging("SMTCLIDStatusChanged", "", "EventID: " & EventID & ", CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                SMTCLIDStatusChanged = "N"
            End Try
        End Using
    End Function
    Public Function SMTCLIDStatusChangedByPC(ByRef EventID As String, ByRef CLID As String, ByRef Action As String, ByRef ResultFlag As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_SMTCycleCount_ChangedStatus_PC"
                sda.SelectCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 150).Value = OracleLoginData.OrgCode
                sda.SelectCommand.Parameters.Add("@EventID", SqlDbType.VarChar, 150).Value = EventID
                sda.SelectCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 150).Value = CLID
                sda.SelectCommand.Parameters.Add("@Action", SqlDbType.VarChar, 150).Value = Action
                sda.SelectCommand.Parameters.Add("@User", SqlDbType.VarChar, 150).Value = OracleLoginData.User
                sda.SelectCommand.Parameters.Add("@Result", SqlDbType.VarChar, 150)
                sda.SelectCommand.Parameters("@Result").Direction = ParameterDirection.Output
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                ResultFlag = sda.SelectCommand.Parameters("@Result").Value
                SMTCLIDStatusChangedByPC = ResultFlag
            Catch ex As Exception
                ErrorLogging("SMTCLIDStatusChanged", "", "EventID: " & EventID & ", CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                SMTCLIDStatusChangedByPC = "N"
            End Try
        End Using
    End Function
#End Region

    Public Function GetCyDate(ByVal cc_name As String, ByVal seq As Integer, ByVal OracleLoginData As ERPLogin) As GetCycleCount
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("CycleList")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_invcc_pkg.get_count_items"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_cc_name", OracleType.VarChar, 50)).Value = cc_name
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_seq", OracleType.Int32)).Value = seq

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "CycleList")
                oda.SelectCommand.Connection.Close()

                GetCyDate.Flag = oda.SelectCommand.Parameters("o_success_flag").Value.ToString.ToUpper()
                GetCyDate.Message = oda.SelectCommand.Parameters("o_error_mssg").Value.ToString

                GetCyDate.ds = ds
            Catch ex As Exception
                ErrorLogging("CycleCount-GetCyDate", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function


    ' ''Public Function PostCycle(ByVal cc_name As String, ByVal seq As Integer) As DataSet

    ' ''    Using da As DataAccess = GetDataAccess()
    ' ''        Try
    ' ''            Dim comm As OracleCommand = da.Ora_Command_Trans()
    ' ''            Dim aa As OracleString
    ' ''            Dim resp As Integer = 53485
    ' ''            Dim resp_appl As Integer = 401

    ' ''            comm.CommandType = CommandType.StoredProcedure
    ' ''            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
    ' ''            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = OracleLoginData.UserID
    ' ''            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
    ' ''            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = resp_appl
    ' ''            comm.ExecuteOracleNonQuery(aa)
    ' ''            comm.Parameters.Clear()


    ' ''            Dim ds As DataSet = New DataSet()
    ' ''            ds.Tables.Add("PostCycleList")
    ' ''            Dim oda As OracleDataAdapter = da.Oda_Sele()
    ' ''            oda.SelectCommand.CommandType = CommandType.StoredProcedure
    ' ''            oda.SelectCommand.CommandText = "apps.xxetr_invcc_pkg.process_cc_count"

    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_cc_name", OracleType.VarChar, 20)).Value = "EMR-NOV09-Counting"
    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = "404"
    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_sequence", OracleType.Int32)).Value = 7
    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_cc_qty", OracleType.Int32)).Value = 50
    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_uom_code", OracleType.VarChar, 20)).Value = "EA"
    ' ''            oda.SelectCommand.Parameters.Add(New OracleParameter("p_reason", OracleType.VarChar, 20)).Value = ""

    ' ''            oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
    ' ''            oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

    ' ''            oda.SelectCommand.Connection.Open()
    ' ''            oda.Fill(ds, "PostCycleList")
    ' ''            oda.SelectCommand.Connection.Close()

    ' ''            Dim flag As String
    ' ''            Dim mess As String

    ' ''            flag = oda.SelectCommand.Parameters("o_success_flag").Value.ToString.ToUpper()
    ' ''            mess = oda.SelectCommand.Parameters("o_error_mssg").Value.ToString()

    ' ''            PostCycle = ds
    ' ''        Catch ex As Exception
    ' ''            ErrorLogging("CycleCount-PostCycle", "", ex.Message & ex.Source, "E")
    ' ''            Throw ex
    ' ''        Finally
    ' ''            If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()    
    ' ''        End Try

    ' ''    End Using
    ' ''End Function


    Public Function UpdateCycle(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString

        Dim resp As Integer = OracleLoginData.RespID_Inv
        Dim appl As Integer = OracleLoginData.AppID_Inv

        Try
            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = OracleLoginData.UserID
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl

            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_invcc_pkg.process_cc_count"
            comm.Parameters.Add("p_cc_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_sequence", OracleType.Int32)
            comm.Parameters.Add("p_cc_qty", OracleType.Double)
            comm.Parameters.Add("p_uom_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
            comm.Parameters.Add("p_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            comm.Parameters("p_cc_name").SourceColumn = "p_cc_name"
            comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("p_sequence").SourceColumn = "p_sequence"
            comm.Parameters("p_cc_qty").SourceColumn = "p_cc_qty"
            comm.Parameters("p_uom_code").SourceColumn = "p_uom_code"
            comm.Parameters("p_reason").SourceColumn = "p_reason"
            comm.Parameters("p_reference").SourceColumn = "p_reference"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("PostList"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("PostList").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging("CycleCount" & "-UpdateCycle", OracleLoginData.User, DR(i)("o_error_mssg").ToString, "I")
                Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If
        Catch ex As Exception
            p_ds.Tables("PostList").Rows(0)("o_success_flag") = "N"
            p_ds.Tables("PostList").Rows(0)("o_error_mssg") = ex.Message
            ErrorLogging("CycleCount" & "-UpdateCycle", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function


    Public Function GetCLIDate(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet

        GetCLIDate = New DataSet
        Dim GetCLIDTab As DataTable
        Dim myDataColumn As DataColumn

        GetCLIDTab = New Data.DataTable("CLIDTab")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QTY", System.Type.GetType("System.Decimal"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNO", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialRevision", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubInv", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Loc", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)

        GetCLIDate.Tables.Add(GetCLIDTab)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select CLID,QtyBaseUOM,MaterialNO,MaterialRevision,SLOC,StorageBin,RTLot,StatusCode from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode", myConn)
        TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        TDHeaderSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20, "OrgCode")
        TDHeaderSQLCommand.Parameters("@CLID").Value = CLID
        TDHeaderSQLCommand.Parameters("@OrgCode").Value = OracleLoginData.OrgCode


        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M5
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = GetCLIDTab.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("QTY") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNO") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MaterialRevision") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("SubInv") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("Loc") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then GetOneRow("RTLot") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then GetOneRow("StatusCode") = objReader.GetValue(7)
                GetCLIDTab.Rows.Add(GetOneRow)
            End While

            myConn.Close()
        Catch ex As Exception
            ErrorLogging("CycleCount" & "-GetCLIDate", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function


    Public Function TBoxData(ByVal BoxID As String, ByVal OracleLoginData As ERPLogin) As DataSet

        TBoxData = New DataSet
        Dim GetCLIDTab As DataTable
        Dim myDataColumn As DataColumn

        GetCLIDTab = New Data.DataTable("BoxIDTab")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QTY", System.Type.GetType("System.Decimal"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNO", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialRevision", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubInv", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Loc", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)

        TBoxData.Tables.Add(GetCLIDTab)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select CLID,QtyBaseUOM,MaterialNO,MaterialRevision,SLOC,StorageBin,RTLot,StatusCode,OrgCode from T_CLMaster where BoxID = @BoxID", myConn)
        TDHeaderSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
        TDHeaderSQLCommand.Parameters("@BoxID").Value = BoxID


        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M5
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = GetCLIDTab.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("QTY") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNO") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MaterialRevision") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("SubInv") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("Loc") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then GetOneRow("RTLot") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then GetOneRow("StatusCode") = objReader.GetValue(7)
                If Not objReader.GetValue(8) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(8)
                GetCLIDTab.Rows.Add(GetOneRow)
            End While

            myConn.Close()
        Catch ex As Exception
            ErrorLogging("CycleCount" & "-TBoxData", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Testupdata(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString

        Dim resp As Integer = OracleLoginData.RespID_Inv
        Dim appl As Integer = OracleLoginData.AppID_Inv

        Try
            'comm.CommandType = CommandType.StoredProcedure
            'comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            'comm.Parameters.Add("p_user_id", OracleType.Int32).Value = OracleLoginData.UserID
            'comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            'comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl

            'comm.ExecuteOracleNonQuery(aa)
            'comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_invcc_pkg.process_cc_count"
            comm.Parameters.Add("p_cc_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_sequence", OracleType.Int32)
            comm.Parameters.Add("p_cc_qty", OracleType.Double)
            comm.Parameters.Add("p_uom_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 240)

            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            comm.Parameters("p_cc_name").SourceColumn = "p_cc_name"
            comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("p_sequence").SourceColumn = "p_sequence"
            comm.Parameters("p_cc_qty").SourceColumn = "p_cc_qty"
            comm.Parameters("p_uom_code").SourceColumn = "p_uom_code"
            comm.Parameters("p_reason").SourceColumn = "p_reason"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("PostList"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("PostList").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging("CycleCount" & "-UpdateCycle", OracleLoginData.User, DR(i)("o_error_mssg").ToString, "I")
                Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If
        Catch ex As Exception
            p_ds.Tables("PostList").Rows(0)("o_success_flag") = "N"
            p_ds.Tables("PostList").Rows(0)("o_error_mssg") = ex.Message
            ErrorLogging("CycleCount" & "-UpdateCycle", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function UpdateReason(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString

        Dim resp As Integer = OracleLoginData.RespID_Inv
        Dim appl As Integer = OracleLoginData.AppID_Inv

        Try
            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = OracleLoginData.UserID
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl

            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_invcc_pkg.update_reason_code"
            comm.Parameters.Add("p_cc_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_sequence", OracleType.Int32)
            comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 240)

            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            comm.Parameters("p_cc_name").SourceColumn = "p_cc_name"
            comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("p_sequence").SourceColumn = "p_sequence"
            comm.Parameters("p_reason").SourceColumn = "p_reason"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("PostList"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("PostList").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging("CycleCount" & "-UpdateReason", OracleLoginData.User, DR(i)("o_error_mssg").ToString, "I")
                Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If
        Catch ex As Exception
            p_ds.Tables("PostList").Rows(0)("o_success_flag") = "N"
            p_ds.Tables("PostList").Rows(0)("o_error_mssg") = ex.Message
            ErrorLogging("CycleCount" & "-UpdateReason", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    ''''''''''''Testing IQC and Putaway report 2010-9-17'''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function IQCPutawayReport(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("IQCPutaway")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_undelivery_pkg.get_iqc_data"

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_po_number", OracleType.VarChar, 20)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_vendor_id", OracleType.Int32)).Value = 0
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_shipnum_fm", OracleType.VarChar, 20)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_shipnum_to", OracleType.VarChar, 50)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_date_fm", OracleType.VarChar, 50)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_date_to", OracleType.VarChar, 20)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_rec_type", OracleType.VarChar, 50)).Value = ""
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_num", OracleType.VarChar, 50)).Value = ""

                oda.SelectCommand.Parameters.Add("o_iqc_data", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "IQCPutaway")
                oda.SelectCommand.Connection.Close()

                IQCPutawayReport = ds

            Catch ex As Exception
                ErrorLogging("IQCPutawayReport", OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function




End Class
