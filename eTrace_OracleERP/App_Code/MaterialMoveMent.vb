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
Imports System.Xml
Imports System.Xml.XmlDocument

Public Structure Result
    Public OracleData As DataSet
    Public Result_Flag As String
    Public ErrMsg As String
End Structure

Public Structure PostBatchRslt
    Public BatchList As DataSet
    Public PrintFlag As Boolean
End Structure

Public Structure BJ_Rs
    Public BJInfo As DataSet
    Public Flag As String
    Public ErrMsg As String
End Structure

Public Structure HW_ExportDataInfo
    Public TransSourceID As String
    Public BarCode As String
    Public SnNO As String
    Public Comments As String
    Public CreatedDate As String
    Public CreatedBy As String
    Public Updated_by As String
    Public Updated_date As String
    Public Segment1 As String
    Public Segment2 As String
    Public Segment3 As String
    Public Segment4 As String
    Public Segment5 As String
    Public Segment6 As String
    Public Segment7 As String
    Public Segment8 As String
    Public Segment9 As String
    Public ErrorCode As String
    Public ErrorMsg As String
End Structure

Public Class MaterialMoveMent
    Inherits PublicFunction
    'Public Function getdataaccess()
    '    Return New DataAccess()
    'End Function

    Public Function UpdateSlotCheckOption(ByVal Options As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                If FixNull(Options) <> "" Then
                    da.ExecuteNonQuery(String.Format("update T_Config set Value='{0}' where ConfigID='HHF018'", Options))
                    Return ""
                End If
            Catch ex As Exception
                ErrorLogging("UpdateSlotCheckOption", OracleLoginData.User, ex.Message & ex.Source, "E")
                Return "Failed"
            End Try
        End Using
    End Function

    Public Function PackingManagement(ByVal FirstScan As String, ByVal CLIDBoxID As String, ByVal PalletWeight As Decimal, ByVal ActionType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_PackingManagement"
                sda.SelectCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 10).Value = OracleLoginData.OrgCode
                sda.SelectCommand.Parameters.Add("@FirstScan", SqlDbType.NVarChar, 100).Value = FirstScan
                sda.SelectCommand.Parameters.Add("@CLIDBoxID", SqlDbType.NVarChar, 100).Value = CLIDBoxID
                sda.SelectCommand.Parameters.Add("@PalletWeight", SqlDbType.Decimal).Value = PalletWeight
                sda.SelectCommand.Parameters.Add("@ActionType", SqlDbType.NVarChar, 100).Value = ActionType
                sda.SelectCommand.Parameters.Add("@User", SqlDbType.NVarChar, 100).Value = OracleLoginData.User
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("PackingManagement", "", "FirstScan: " & FirstScan & ", " & "CLIDBoxID: " & CLIDBoxID & ", " & "ActionType: " & ActionType & ", " & ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function SFGetDcodeLnforIntSN(ByVal IntSN As String, ByVal PCBA As String, ByVal Component As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql, InformList As String
            Dim i As Integer
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_SFGetDcodeLnforIntSN '{0}','{1}','{2}'", IntSN, PCBA, Component)
                ds = da.ExecuteDataSet(sql)

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        InformList = InformList & ds.Tables(0).Rows(i)(0) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(1) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(2) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(3) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(4) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(5) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(6) & ";"
                    Next
                Else
                    SFGetDcodeLnforIntSN = ""
                End If
                SFGetDcodeLnforIntSN = FixNull(InformList)
            Catch ex As Exception
                ErrorLogging("SFGetDcodeLnforIntSN", "", "IntSN " & IntSN & ", " & "PCBA " & PCBA & ", " & "Component " & Component & ", " & ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function SFGetDcodeLnforSN(ByVal SN As String, ByVal PCBA As String, ByVal Component As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql, InformList As String
            Dim i As Integer
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_SFGetDcodeLnforSN '{0}','{1}','{2}'", SN, PCBA, Component)
                ds = da.ExecuteDataSet(sql)

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        InformList = InformList & ds.Tables(0).Rows(i)(0) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(1) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(2) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(3) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(4) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(5) & ">"
                        InformList = InformList & ds.Tables(0).Rows(i)(6) & ";"
                    Next
                Else
                    SFGetDcodeLnforSN = ""
                End If
                SFGetDcodeLnforSN = FixNull(InformList)
            Catch ex As Exception
                ErrorLogging("SFGetDcodeLnforSN", "", "SN " & SN & ", " & "PCBA " & PCBA & ", " & "Component " & Component & ", " & ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function transfer_submit(ByVal p_is_return_rcv_no As String, ByVal p_timeout As String, ByVal p_transaction_header_id As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim Oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("MSG")
                Dim comm_submit As OracleCommand = da.OraCommand()
                Oda.InsertCommand.CommandType = CommandType.StoredProcedure
                Oda.InsertCommand.CommandText = "APPS.emrsn_Mtl_Tran_Process_OnLine.Call_Process_OnLine"
                Oda.InsertCommand.Parameters.Add("o_error_cursor", OracleType.Cursor).Direction = ParameterDirection.Output

                Oda.InsertCommand.Parameters.Add("p_transaction_header_id", OracleType.Int32).Value = p_transaction_header_id
                Oda.InsertCommand.Parameters.Add("p_timeout", OracleType.Int32).Value = p_timeout

                Oda.InsertCommand.Parameters.Add("p_error_code", OracleType.VarChar, 2)
                Oda.InsertCommand.Parameters.Add("p_error_explanation", OracleType.VarChar, 2)
                Oda.InsertCommand.Parameters.Add("p_process_online_yn", OracleType.Blob)
                Oda.InsertCommand.Parameters.Add("x_return_status", OracleType.VarChar, 2)
                Oda.InsertCommand.Parameters.Add("x_msg_count", OracleType.Int32)
                Oda.InsertCommand.Parameters.Add("x_msg_data", OracleType.VarChar, 2)
                Oda.InsertCommand.Parameters.Add("x_trans_count", OracleType.Int32)
                Oda.InsertCommand.Parameters.Add("p_tran_result", OracleType.Int32)

                Oda.InsertCommand.Parameters("p_error_code").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("p_error_explanation").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("p_process_online_yn").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("x_return_status").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("x_msg_count").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("x_msg_data").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("x_trans_count").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("p_tran_result").Direction = ParameterDirection.Output

                Oda.InsertCommand.Connection.Open()
                Oda.Fill(ds, "MSG")
                Oda.SelectCommand.Connection.Close()
                'Oda.Dispose()
                Return ds
            Catch oe As OracleException
                ErrorLogging("MatMovement-mat_submit_ds", "Yong", oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function MatSourceRead(ByVal CLID As String, ByVal OrgCode As String, ByVal SourceSubInv As String, ByVal SourceLocator As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim i As Integer
        MatSourceRead = New DataSet
        Dim SourceData As DataTable
        Dim myDataColumn As DataColumn

        Try
            SourceData = New Data.DataTable("SourceData")  'Add Table

            myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MaterialRevision", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Double"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("BaseUOM", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("SourceBin", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("TotalQty", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("SubRc", System.Type.GetType("System.Int16"))
            SourceData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ErrMSG", System.Type.GetType("System.String"))
            SourceData.Columns.Add(myDataColumn)

            Dim ReadDataSQLCommand As SqlClient.SqlCommand
            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
            Dim objReader As SqlClient.SqlDataReader
            CLID = Trim(CLID)
            If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                If SourceSubInv = "" And SourceLocator = "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then  ' and not(LastDJ = '' or LastDJ IS NULL)
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode in ('0', '1') and SLOC like '%OS%' and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode = '1' And SLOC Like '%MRB%'", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode = '1'", myConn) '
                    Else
                        'For UAT, the CLID status can be "1" or "0" while doing SubInv Transfer
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                    ReadDataSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
                    ReadDataSQLCommand.Parameters("@BoxID").Value = CLID
                ElseIf SourceSubInv <> "" And SourceLocator = "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and StatusCode in ('0', '1', '5') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and SLOC like '%OS%' and StatusCode in ('0', '1') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StatusCode = '1' and SLOC like '%MRB'", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StatusCode = '1'", myConn) '
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                    ReadDataSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
                    ReadDataSQLCommand.Parameters("@BoxID").Value = CLID
                    ReadDataSQLCommand.Parameters.Add("@SourceSubInv", SqlDbType.VarChar, 50, "SourceSubInv")
                    ReadDataSQLCommand.Parameters("@SourceSubInv").Value = SourceSubInv
                ElseIf SourceSubInv <> "" And SourceLocator <> "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and StorageBin = @SourceLocator and StatusCode in ('0', '1', '5') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and SLOC like '%OS%' and StorageBin = @SourceLocator and StatusCode in ('0', '1') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1' AND SLOC like '%MRB'", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1' ", myConn) '
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and BoxID = @BoxID and SLOC = @SourceSubInv and StorageBin = @SourceLocator and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                    ReadDataSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
                    ReadDataSQLCommand.Parameters("@BoxID").Value = CLID
                    ReadDataSQLCommand.Parameters.Add("@SourceSubInv", SqlDbType.VarChar, 50, "SourceSubInv")
                    ReadDataSQLCommand.Parameters("@SourceSubInv").Value = SourceSubInv
                    ReadDataSQLCommand.Parameters.Add("@SourceLocator", SqlDbType.VarChar, 20, "SourceLocator")
                    ReadDataSQLCommand.Parameters("@SourceLocator").Value = SourceLocator
                End If

            Else 'normal CLID
                If SourceSubInv = "" And SourceLocator = "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0', '1') and SLOC like '%OS%' and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        'ErrorLogging(MoveType & "-SourceRead", "ANDYKING_TANG", "MRB", "I")
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode = '1' and SLOC like '%MRB%'", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode = '1' ", myConn) '
                    Else
                        'ErrorLogging(MoveType & "-SourceRead", "ANDYKING_TANG", "OTHERS", "I")
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
                    ReadDataSQLCommand.Parameters("@CLID").Value = CLID
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                ElseIf SourceSubInv <> "" And SourceLocator = "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and StatusCode IN ('0', '1', '5') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and SLOC like '%OS%' and StatusCode IN ('0', '1') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StatusCode = '1' and SLOC like '%MRB%'", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StatusCode = '1'", myConn) '
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
                    ReadDataSQLCommand.Parameters("@CLID").Value = CLID
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                    ReadDataSQLCommand.Parameters.Add("@SourceSubInv", SqlDbType.VarChar, 50, "SourceSubInv")
                    ReadDataSQLCommand.Parameters("@SourceSubInv").Value = SourceSubInv
                ElseIf SourceSubInv <> "" And SourceLocator <> "" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn) '
                    ElseIf MoveType = "MMt-FSRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and StorageBin = @SourceLocator and StatusCode in ('0', '1', '5') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and SLOC like '%OS%' and StorageBin = @SourceLocator and StatusCode in ('0', '1') and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn) '
                    ElseIf MoveType = "MMt-MRBRTV" Or MoveType = "MMt-MRBSCR" Or MoveType = "MMt-MRBRWK" Or MoveType = "MMt-MRBSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1' and SLOC like '%MRB%", myConn) '
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1' ", myConn) '
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode = '1'", myConn) '
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and StatusCode in ('0', '1')", myConn) '
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID,MaterialNo ,MaterialRevision,QtyBaseUOM,BaseUOM,RTLot,SLOC,StorageBin from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and SLOC = @SourceSubInv and StorageBin = @SourceLocator and (StatusCode = '1' or StatusCode = '4')", myConn) '
                    End If
                    ReadDataSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
                    ReadDataSQLCommand.Parameters("@CLID").Value = CLID
                    ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                    ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                    ReadDataSQLCommand.Parameters.Add("@SourceSubInv", SqlDbType.VarChar, 50, "SourceSubInv")
                    ReadDataSQLCommand.Parameters("@SourceSubInv").Value = SourceSubInv
                    ReadDataSQLCommand.Parameters.Add("@SourceLocator", SqlDbType.VarChar, 20, "SourceLocator")
                    ReadDataSQLCommand.Parameters("@SourceLocator").Value = SourceLocator
                End If

            End If

            Dim SourceDataRow As Data.DataRow
            SourceDataRow = SourceData.NewRow()
            Dim L_TotalQty As Double
            Dim L_MaterialNo As String = ""
            Dim L_MaterialRV As String = ""
            Dim L_RTLot As String = ""
            Dim L_SLOC As String = ""
            Dim L_StorageBin As String = ""
            Dim flag As Integer
            flag = 0
            Try
                myConn.Open()
                ReadDataSQLCommand.CommandTimeout = TimeOut_M5
                objReader = ReadDataSQLCommand.ExecuteReader()
                While objReader.Read()
                    If flag = 0 Then
                        MatSourceRead.Tables.Add(SourceData)
                    End If
                    flag = 1

                    If Not objReader.GetValue(1) Is DBNull.Value Then         'MaterialNo
                        SourceDataRow("MaterialNo") = objReader.GetValue(1)
                        L_MaterialNo = objReader.GetValue(1)
                    End If
                    If Not objReader.GetValue(2) Is DBNull.Value Then         'MaterialNo
                        SourceDataRow("MaterialRevision") = objReader.GetValue(2)
                        L_MaterialRV = objReader.GetValue(2)
                    Else
                        SourceDataRow("MaterialRevision") = ""
                    End If

                    If Not objReader.GetValue(3) Is DBNull.Value Then
                        SourceDataRow("QtyBaseUOM") = objReader.GetValue(3) '
                    Else
                        SourceDataRow("QtyBaseUOM") = ""
                    End If

                    If Not objReader.GetValue(4) Is DBNull.Value Then
                        SourceDataRow("BaseUOM") = objReader.GetValue(4)
                    Else
                        SourceDataRow("BaseUOM") = ""
                    End If

                    If Not objReader.GetValue(5) Is DBNull.Value Then              'SourceBin
                        SourceDataRow("RTLot") = objReader.GetValue(5)
                        L_RTLot = objReader.GetValue(5)
                    Else
                        SourceDataRow("RTLot") = ""
                    End If

                    If Not objReader.GetValue(6) Is DBNull.Value Then              'SourceBin
                        SourceDataRow("SLOC") = objReader.GetValue(6)
                        L_SLOC = objReader.GetValue(6)
                    Else
                        SourceDataRow("SLOC") = ""
                    End If

                    If Not objReader.GetValue(7) Is DBNull.Value Then              'SourceBin
                        SourceDataRow("SourceBin") = objReader.GetValue(7)
                        L_StorageBin = objReader.GetValue(7)
                    Else
                        SourceDataRow("SourceBin") = ""
                    End If

                    Exit While
                End While
                objReader.Close()
                myConn.Close()
            Catch ex As Exception
                ErrorLogging(MoveType & "-SourceRead", OracleLoginData.User, ex.Message & ex.Source, "E")
                Exit Function
            Finally
                If myConn.State <> ConnectionState.Closed Then myConn.Close()
            End Try

            If flag <> 0 Then
                SourceData.Rows.Add(SourceDataRow)
                Dim CLIDS As DataTable
                CLIDS = New Data.DataTable("CLIDS")
                myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Double"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.DateTime"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("StorageBin", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("BoxID", System.Type.GetType("System.String"))
                CLIDS.Columns.Add(myDataColumn)
                If MoveType = "MMt-ProdRtn" Or MoveType = "MMt-OSPRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-CostSub" Then
                    myDataColumn = New Data.DataColumn("SupplyType", System.Type.GetType("System.String"))
                    CLIDS.Columns.Add(myDataColumn)
                    myDataColumn = New Data.DataColumn("LastDJ", System.Type.GetType("System.String"))
                    CLIDS.Columns.Add(myDataColumn)
                End If
                If MoveType = "MMt-ProdRtn" Then
                    myDataColumn = New Data.DataColumn("MSL", System.Type.GetType("System.String"))
                    CLIDS.Columns.Add(myDataColumn)
                    myDataColumn = New Data.DataColumn("SMCLID", System.Type.GetType("System.String"))
                    CLIDS.Columns.Add(myDataColumn)
                    myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
                    CLIDS.Columns.Add(myDataColumn)
                End If
                myDataColumn = New Data.DataColumn("Flag", System.Type.GetType("System.Boolean"))
                CLIDS.Columns.Add(myDataColumn)
                MatSourceRead.Tables.Add(CLIDS)

                'Check the transfer logic
                MatSourceRead.Tables("SourceData").Rows(0)("SubRc") = 0

                Dim Check_Lot As String
                Check_Lot = ""
                If FixNull(DestSubInv) <> "" AndAlso InStr(FixNull(DestSubInv), "OS") < 1 AndAlso InStr(FixNull(DestSubInv), "FS") < 1 AndAlso MoveType <> "MMt-SubCost" Then
                    Using da As DataAccess = GetDataAccess()
                        Try
                            Check_Lot = Convert.ToString(da.ExecuteScalar(String.Format("select Value from T_Config with (nolock) where ConfigID = 'MMT001'")))
                            If Not Check_Lot Is Nothing AndAlso Not Check_Lot Is DBNull.Value AndAlso Check_Lot.ToString = "YES" Then
                                Dim Sqlstr As String
                                Dim checkdata As DataSet = New DataSet
                                Sqlstr = String.Format("SELECT TOP (1) CLID, SLOC, StorageBin, MaterialNo, RTLot  FROM T_CLMaster with (nolock) WHERE  StatusCode = '1' and MaterialNo = '{0}' and RTLot <> '{1}' and SLOC = '{2}' and StorageBin = '{3}' and OrgCode = '{4}'", Trim(L_MaterialNo), Trim(L_RTLot), FixNull(DestSubInv), FixNull(DestLocator), OrgCode)
                                checkdata = da.ExecuteDataSet(Sqlstr, "DestData")
                                If Not checkdata Is Nothing AndAlso checkdata.Tables.Count > 0 AndAlso checkdata.Tables(0).Rows.Count > 0 Then
                                    MatSourceRead.Tables("SourceData").Rows(0)("SubRc") = 2
                                    MatSourceRead.Tables("SourceData").Rows(0)("ErrMSG") = "Different LotNo exists in the same locator. Pls transfer to another locator!"
                                End If
                            End If
                        Catch ex As Exception
                            Throw ex
                            ErrorLogging(MoveType & "CheckDest", OracleLoginData.User, ex.Message & ex.Source, "E")
                        End Try
                    End Using
                End If

                Dim CLIDSRow As Data.DataRow
                'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM from T_CLMaster where RecDocNo = @RecDocNo and RecDocYear = @RecDocYear and MaterialNo = @MaterialNo and StorageType = @StorageType", myConn)
                If FixNull(Check_Lot) = "" OrElse FixNull(Check_Lot).ToUpper = "NO" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn)
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select T_CLMaster.CLID, QtyBaseUOM, BoxID , StatusCode, T_CLMaster.ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ,T_CLMaster.MSL, T_SMCLID.CLID as SMCLID, T_SMCLID.Status from T_CLMaster with (nolock) left outer join T_SMCLID with (nolock) on T_CLMaster.CLID = T_SMCLID.CLID where OrgCode = @OrgCode and (StatusCode = '0' or StatusCode = '5') and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and T_CLMaster.LastTransaction not like 'BF Split%' and T_CLMaster.LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select T_CLMaster.CLID, QtyBaseUOM, BoxID , StatusCode, T_CLMaster.ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ,T_CLMaster.MSL, T_SMCLID.CLID as SMCLID, T_SMCLID.Status from T_CLMaster with (nolock) left outer join T_SMCLID with (nolock) on T_CLMaster.CLID = T_SMCLID.CLID where T_CLMaster.CLID = @CLID and OrgCode = @OrgCode and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and T_CLMaster.LastTransaction not like 'BF Split%' and T_CLMaster.LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-FSRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock)))' and BoxID = @BoxID and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and StatusCode in ('0', '1') and SLOC like '%OS%' and BoxID = @BoxID and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0','1') and SLOC like '%OS%' and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode = '1'", myConn)
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin from T_CLMaster where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode = '1'", myConn)
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin from T_CLMaster where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode in ('0', '1')", myConn)
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and (StatusCode = '1' or StatusCode = '4')", myConn)
                    End If

                ElseIf Check_Lot.ToUpper = "YES" Then
                    If MoveType = "MMt-CostSub" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and RTLot = @RTLot and StatusCode = '0' and (SupplyType = '' or SupplyType IS NULL) and (LastDJ = '' or LastDJ IS NULL)", myConn)
                    ElseIf MoveType = "MMt-ProdRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select T_CLMaster.CLID, QtyBaseUOM, BoxID , StatusCode, T_CLMaster.ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ,T_CLMaster.MSL, T_SMCLID.CLID as SMCLID, T_SMCLID.Status from T_CLMaster with (nolock) left outer join T_SMCLID with (nolock) on T_CLMaster.CLID = T_SMCLID.CLID where OrgCode = @OrgCode and (StatusCode = '0' or StatusCode = '5') and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and T_CLMaster.LastTransaction not like 'BF Split%' and T_CLMaster.LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select T_CLMaster.CLID, QtyBaseUOM, BoxID , StatusCode, T_CLMaster.ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ,T_CLMaster.MSL, T_SMCLID.CLID as SMCLID, T_SMCLID.Status from T_CLMaster with (nolock) left outer join T_SMCLID with (nolock) on T_CLMaster.CLID = T_SMCLID.CLID where T_CLMaster.CLID = @CLID and OrgCode = @OrgCode and (StatusCode = '0' or StatusCode = '5') and not(SupplyType = '' or SupplyType IS NULL) and T_CLMaster.LastTransaction not like 'BF Split%' and T_CLMaster.LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-FSRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock)))' and BoxID = @BoxID and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0', '1', '5') and (SLOC like '%FS%' or SLOC in (select Subinventory from T_BackflushSubinv with (nolock))) and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-OSPRtn" Then
                        If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where OrgCode = @OrgCode and StatusCode in ('0', '1') and SLOC like '%OS%' and BoxID = @BoxID and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and BoxID = @BoxID and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        Else
                            ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin, SupplyType, LastDJ from T_CLMaster with (nolock) where CLID = @CLID and OrgCode = @OrgCode and StatusCode in ('0','1') and SLOC like '%OS%' and LastTransaction not like 'BF Split%' and LastTransaction not like 'CLID Split%'", myConn)
                            'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin, SupplyType, LastDJ from T_CLMaster where CLID = @CLID and OrgCode = @OrgCode and ((StatusCode = '0' and SupplyType <> 'OSP') or (StatusCode in ('0','1') and SupplyType = 'OSP')) and not(SupplyType = '' or SupplyType IS NULL) and not(LastDJ = '' or LastDJ IS NULL)", myConn)
                        End If
                    ElseIf MoveType = "MMt-SubCost" Then
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and RTLot = @RTLot and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode = '1'", myConn)
                    Else
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin from T_CLMaster where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode = '1'", myConn)
                        'ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , ExpDate, RTLot, StorageBin from T_CLMaster where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and SLOC =@SLOC and StorageBin = @StorageBin and StatusCode in ('0', '1')", myConn)
                        ReadDataSQLCommand = New SqlClient.SqlCommand("Select CLID, QtyBaseUOM, BoxID , StatusCode, ExpDate, RTLot, SLOC, StorageBin from T_CLMaster with (nolock) where OrgCode = @OrgCode and MaterialNo = @MaterialNo and MaterialRevision = @MaterialRevision and RTLot = @RTLot and SLOC =@SLOC and StorageBin = @StorageBin and (StatusCode = '1' or StatusCode = '4')", myConn)
                    End If
                End If

                If (MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-OSPRtn") And ((Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P") Then
                    ReadDataSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 50, "BoxID")
                ElseIf (MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-OSPRtn") And Not ((Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P") Then
                    ReadDataSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 50, "CLID")
                End If
                ReadDataSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
                If Not MoveType = "MMt-ProdRtn" And Not MoveType = "MMt-FSRtn" And Not MoveType = "MMt-OSPRtn" Then
                    ReadDataSQLCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 50, "MaterialNo")
                    ReadDataSQLCommand.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 50, "MaterialRevision")
                End If
                If Not MoveType = "MMt-CostSub" And Not MoveType = "MMt-ProdRtn" And Not MoveType = "MMt-FSRtn" And Not MoveType = "MMt-OSPRtn" Then
                    ReadDataSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 20, "SLOC")
                    ReadDataSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
                End If
                If Check_Lot = "YES" AndAlso (MoveType = "MMt-CostSub" Or MoveType = "MMt-SubTran") Then
                    ReadDataSQLCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 20, "RTLot")
                End If


                If (MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-OSPRtn") And ((Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P") Then
                    ReadDataSQLCommand.Parameters("@BoxID").Value = CLID
                ElseIf (MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-OSPRtn") And Not ((Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P") Then
                    ReadDataSQLCommand.Parameters("@CLID").Value = CLID
                End If
                ReadDataSQLCommand.Parameters("@OrgCode").Value = OrgCode
                If Not MoveType = "MMt-ProdRtn" And Not MoveType = "MMt-FSRtn" And Not MoveType = "MMt-OSPRtn" Then
                    ReadDataSQLCommand.Parameters("@MaterialNo").Value = Trim(L_MaterialNo)
                    ReadDataSQLCommand.Parameters("@MaterialRevision").Value = Trim(L_MaterialRV)
                End If
                If Not MoveType = "MMt-CostSub" And Not MoveType = "MMt-FSRtn" And Not MoveType = "MMt-ProdRtn" And Not MoveType = "MMt-OSPRtn" Then
                    ReadDataSQLCommand.Parameters("@SLOC").Value = Trim(L_SLOC)
                    ReadDataSQLCommand.Parameters("@StorageBin").Value = Trim(L_StorageBin)
                End If
                If Check_Lot = "YES" AndAlso (MoveType = "MMt-CostSub" Or MoveType = "MMt-SubTran") Then
                    ReadDataSQLCommand.Parameters("@RTLot").Value = Trim(L_RTLot)
                End If


                Try
                    myConn.Open()
                    ReadDataSQLCommand.CommandTimeout = TimeOut_M5
                    objReader = ReadDataSQLCommand.ExecuteReader()
                    While objReader.Read()
                        CLIDSRow = CLIDS.NewRow()
                        If Not objReader.GetValue(0) Is DBNull.Value Then CLIDSRow("CLID") = objReader.GetValue(0).ToString
                        If Not objReader.GetValue(1).ToString Is DBNull.Value Then
                            CLIDSRow("QtyBaseUOM") = objReader.GetValue(1)
                            L_TotalQty = L_TotalQty + objReader.GetValue(1)
                        End If
                        CLIDSRow("Flag") = False
                        If Not objReader.GetValue(2) Is DBNull.Value Then CLIDSRow("BoxID") = objReader.GetValue(2).ToString
                        If Not objReader.GetValue(3) Is DBNull.Value Then CLIDSRow("StatusCode") = objReader.GetValue(3).ToString
                        If Not objReader.GetValue(4) Is DBNull.Value Then CLIDSRow("ExpDate") = objReader.GetValue(4).ToString
                        If Not objReader.GetValue(5) Is DBNull.Value Then CLIDSRow("RTLot") = objReader.GetValue(5).ToString
                        If Not objReader.GetValue(6) Is DBNull.Value Then CLIDSRow("SLOC") = objReader.GetValue(6).ToString
                        If Not objReader.GetValue(7) Is DBNull.Value Then CLIDSRow("StorageBin") = objReader.GetValue(7).ToString
                        If MoveType = "MMt-CostSub" Or MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Or MoveType = "MMt-OSPRtn" Then
                            If Not objReader.GetValue(8) Is DBNull.Value Then CLIDSRow("SupplyType") = objReader.GetValue(8).ToString
                            If Not objReader.GetValue(9) Is DBNull.Value Then CLIDSRow("LastDJ") = objReader.GetValue(9).ToString
                        End If
                        If MoveType = "MMt-ProdRtn" Then
                            If Not objReader.GetValue(10) Is DBNull.Value Then CLIDSRow("MSL") = objReader.GetValue(10).ToString
                            If Not objReader.GetValue(11) Is DBNull.Value Then CLIDSRow("SMCLID") = objReader.GetValue(11).ToString
                            If Not objReader.GetValue(12) Is DBNull.Value Then CLIDSRow("Status") = objReader.GetValue(12).ToString
                        End If
                        CLIDS.Rows.Add(CLIDSRow)
                    End While
                    objReader.Close()
                    myConn.Close()
                Catch ex As Exception
                    ErrorLogging(MoveType & "-SourceRead", OracleLoginData.User, ex.Message & ex.Source, "E")
                    Exit Function
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try
            End If

            If flag <> 0 Then
                For i = 0 To MatSourceRead.Tables("SourceData").Rows.Count - 1
                    MatSourceRead.Tables("SourceData").Rows(i)("TotalQty") = L_TotalQty
                Next
            End If
        Catch ex As Exception
            ErrorLogging(MoveType & "-SourceRead", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function ClearPalletID(ByVal CartonID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                If CartonID <> "" Then
                    da.ExecuteNonQuery(String.Format("update T_CLMaster set BoxID='' where CLID='{0}'", CartonID))
                    da.ExecuteNonQuery(String.Format("update T_Shippment set PalletID ='' where CartonID='{0}' and (OrgCode='{1}' or OrgCode IS NULL) ", CartonID, OracleLoginData.OrgCode))
                    Return True
                End If
            Catch ex As Exception
                ErrorLogging("ClearPalletID", OracleLoginData.User, ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function CheckCLID(ByVal CLID As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                If (Mid(CLID, 3, 1) = "B" AndAlso Len(CLID) = 20) OrElse Mid(CLID, 1, 1) = "P" Then 'Box ID
                    'Sqlstr = String.Format("SELECT top 1 MaterialNo,MaterialRevision,MaterialDesc,BaseUOM,DateCode, LotNo,RTLot,ExpDate,ROHS,Manufacturer,ManufacturerPN,Stemp,MSL from T_CLMaster where BoxID = '{0}' and OrgCode='{1}'", CLID, OracleLoginData.OrgCode)
                    Sqlstr = String.Format("SELECT top 1 * from T_CLMaster with (nolock) where BoxID = '{0}' and OrgCode='{1}'", CLID, OracleLoginData.OrgCode)
                Else
                    Sqlstr = String.Format("SELECT * from T_CLMaster with (nolock) where CLID = '{0}' and OrgCode='{1}'", CLID, OracleLoginData.OrgCode)
                End If
                Return da.ExecuteDataSet(Sqlstr, "CLIDInfo")

            Catch ex As Exception
                ErrorLogging(MoveType & "CheckCLID", OracleLoginData.User.ToUpper, "CLID: " & CLID & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function RepSourceRead(ByVal CLID As String, ByVal OrgCode As String, ByVal Item As String, ByVal Revision As String, ByVal SourceSubInv As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                    Sqlstr = String.Format("Select CLID, QtyBaseUOM, BaseUOM, ExpDate, RTLot, StorageBin, BoxID from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{2}' and MaterialRevision = '{3}' and SLOC ='{4}' and BoxID = '{0}' and StatusCode = '1'", CLID, OrgCode, Item, Revision, SourceSubInv)
                Else
                    Sqlstr = String.Format("Select CLID, QtyBaseUOM, BaseUOM, ExpDate, RTLot, StorageBin, BoxID from T_CLMaster with (nolock) where CLID = '{0}' and OrgCode = '{1}' and MaterialNo = '{2}' and MaterialRevision = '{3}' and SLOC ='{4}' and StatusCode = '1'", CLID, OrgCode, Item, Revision, SourceSubInv)
                End If
                Return da.ExecuteDataSet(Sqlstr, "CLIDS")
            End Using
        Catch ex As Exception
            ErrorLogging(MoveType & "-RepSourceRead", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
        'Sqlstr = String.Format("Select CLID, QtyBaseUOM, BaseUOM, ExpDate, RTLot, StorageBin, BoxID, 'False' as Flag from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and MaterialRevision = '{2}' and SLOC ='{3}' and StatusCode = '1'", OrgCode, Item, Revision, SourceSubInv)
    End Function

    'Public Function PI_ClidUpdate(ByVal CLIDList As DataSet, ByVal UserName As String) As String
    '    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
    '    Dim SQLTransaction As SqlClient.SqlTransaction
    '    Dim CLMasterSQLCommand As SqlClient.SqlCommand
    '    Dim i, ra As Integer
    '    Dim strCMD As String

    '    Dim CheckData As DataTable
    '    Dim myDataColumn As DataColumn
    '    Dim myDataRow As DataRow
    '    CheckData = New Data.DataTable("CheckData")
    '    myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
    '    CheckData.Columns.Add(myDataColumn)
    '    myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Double"))
    '    CheckData.Columns.Add(myDataColumn)
    '    myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
    '    CheckData.Columns.Add(myDataColumn)
    '    myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
    '    CheckData.Columns.Add(myDataColumn)
    '    myDataColumn = New Data.DataColumn("StorageBin", System.Type.GetType("System.String"))
    '    CheckData.Columns.Add(myDataColumn)
    '    CLIDList.Tables.Add(CheckData)

    '    myDataRow = CLIDList.Tables("CheckData").NewRow()
    '    myDataRow(0) = "FY000000000000000001"
    '    myDataRow(1) = 1
    '    myDataRow(2) = 1
    '    myDataRow(3) = "FR0 MC13"
    '    myDataRow(4) = "TLA1.01.01.01.1"

    '    CLIDList.Tables(0).Rows.Add(myDataRow)
    '    myDataRow = CLIDList.Tables("CheckData").NewRow()
    '    myDataRow(0) = "FY000000000000000002"
    '    myDataRow(1) = 2
    '    myDataRow(2) = 1
    '    myDataRow(3) = "FR0 MC8FLKSAJFDDDDAAAAAAAAAAAAAAAAAAA"
    '    myDataRow(4) = "TLA1.01.21.01.2"
    '    CLIDList.Tables(0).Rows.Add(myDataRow)

    '    myDataRow = CLIDList.Tables("CheckData").NewRow()
    '    myDataRow(0) = "FY000000000000000003"
    '    myDataRow(1) = 3
    '    myDataRow(2) = 1
    '    myDataRow(3) = "FR0 MC1DJFFFFFFFFFFFFFFFFFDKDKDKDKDKD"
    '    myDataRow(4) = "TLA1.01.01.01.3"
    '    CLIDList.Tables(0).Rows.Add(myDataRow)

    '    myDataRow = CLIDList.Tables("CheckData").NewRow()
    '    myDataRow(0) = "FY000000000000000004"
    '    myDataRow(1) = 4
    '    myDataRow(2) = 1
    '    myDataRow(3) = "FR0 MC10"
    '    myDataRow(4) = "TLA1.01.01.01.4"
    '    CLIDList.Tables(0).Rows.Add(myDataRow)

    '    Try
    '        If CLIDList Is Nothing OrElse CLIDList.Tables.Count < 1 OrElse CLIDList.Tables(0).Rows.Count < 1 Then
    '            Return "Error: No record for udpate"
    '        End If
    '        If FixNull(UserName) = "" Then
    '            Return "Error: Username is empty"
    '        End If
    '        myConn.Open()
    '        SQLTransaction = myConn.BeginTransaction
    '        'CLMasterSQLCommand.Connection = myConn
    '        'CLMasterSQLCommand.Transaction = SQLTransaction

    '        For i = 0 To CLIDList.Tables(0).Rows.Count - 1
    '            strCMD = String.Format("UPDATE T_CLMASTER set QtyBaseUOM='{0}',ChangedOn = getdate(), ChangedBy = '{1}', StatusCode = '{2}', SLOC = '{3}', StorageBin = '{4}', LastTransaction = 'PI_ClidUpdate' where CLID = '{5}'", CLIDList.Tables(0).Rows(i)("QtyBaseUOM"), UserName, CLIDList.Tables(0).Rows(i)("StatusCode"), CLIDList.Tables(0).Rows(i)("SLOC"), CLIDList.Tables(0).Rows(i)("StorageBin"), CLIDList.Tables(0).Rows(i)("CLID"))
    '            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
    '            CLMasterSQLCommand.Connection = myConn
    '            CLMasterSQLCommand.Transaction = SQLTransaction
    '            ra = CLMasterSQLCommand.ExecuteNonQuery()
    '        Next
    '        SQLTransaction.Commit()
    '        PI_ClidUpdate = ""
    '    Catch ex As Exception
    '        SQLTransaction.Rollback()
    '        PI_CLIDUpdate = "Error while updating. No record has been udpated."
    '        ErrorLogging("PI_ClidUpdate", UserName, ex.Message & ex.Source)
    '    Finally
    '        If myConn.State <> ConnectionState.Closed Then myConn.Close()
    '    End Try

    'End Function

    Public Function PI_ClidUpdate(ByVal CLIDList As DataSet, ByVal UserName As String) As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String

        Dim CheckData As DataTable
        Dim myDataColumn As DataColumn
        Dim myDataRow As DataRow
        CheckData = New Data.DataTable("CheckData")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CheckData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        CheckData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
        CheckData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StorageBin", System.Type.GetType("System.String"))
        CheckData.Columns.Add(myDataColumn)
        CLIDList.Tables.Add(CheckData)

        myDataRow = CLIDList.Tables("CheckData").NewRow()
        myDataRow(0) = "FY000000000000000001"
        myDataRow(1) = 1
        myDataRow(2) = "FR0 MC13"
        myDataRow(3) = "TLA1.01.01.01.1"

        CLIDList.Tables(0).Rows.Add(myDataRow)
        myDataRow = CLIDList.Tables("CheckData").NewRow()
        myDataRow(0) = "FY000000000000000002"
        myDataRow(1) = 1
        myDataRow(2) = "FR0 MC8FLKSAJFDDDDAAAAAAAAAAAAAAAAAAAjdklsafjlsdakfjsdklfjsdklfjsdkljfksdalf"
        myDataRow(3) = "TLA1.01.21.01.2"
        CLIDList.Tables(0).Rows.Add(myDataRow)

        myDataRow = CLIDList.Tables("CheckData").NewRow()
        myDataRow(0) = "FY000000000000000003"
        myDataRow(1) = 1
        myDataRow(2) = "FR0 MC1DJFFFFFFFFFFFFFFFFFDKDKDKDKDKDfsdofjsdaklfjsdlkrjfweoirueiwjsdjfsdklf"
        myDataRow(3) = "TLA1.01.01.01.3"
        CLIDList.Tables(0).Rows.Add(myDataRow)

        myDataRow = CLIDList.Tables("CheckData").NewRow()
        myDataRow(0) = "FY000000000000000004"
        myDataRow(1) = 1
        myDataRow(2) = "FR0 MC10"
        myDataRow(3) = "TLA1.01.01.01.4"
        CLIDList.Tables(0).Rows.Add(myDataRow)

        Try
            If CLIDList Is Nothing OrElse CLIDList.Tables.Count < 1 OrElse CLIDList.Tables(0).Rows.Count < 1 Then
                Return "Error: No record for udpate"
            End If
            If FixNull(UserName) = "" Then
                Return "Error: Username is empty"
            End If
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction
            'CLMasterSQLCommand.Connection = myConn
            'CLMasterSQLCommand.Transaction = SQLTransaction

            For i = 0 To CLIDList.Tables(0).Rows.Count - 1
                strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = '{1}', SLOC = '{2}', StorageBin = '{3}', LastTransaction = 'PI_ClidUpdate' where CLID = '{4}'", UserName, CLIDList.Tables(0).Rows(i)("StatusCode"), CLIDList.Tables(0).Rows(i)("SLOC"), CLIDList.Tables(0).Rows(i)("StorageBin"), CLIDList.Tables(0).Rows(i)("CLID"))
                CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                CLMasterSQLCommand.Connection = myConn
                CLMasterSQLCommand.Transaction = SQLTransaction
                ra = CLMasterSQLCommand.ExecuteNonQuery()
            Next
            SQLTransaction.Commit()
            PI_ClidUpdate = ""
        Catch ex As Exception
            SQLTransaction.Rollback()
            PI_ClidUpdate = "Error while updating. No record has been udpated."
            ErrorLogging("PI_ClidUpdate", UserName, ex.Message & ex.Source)
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function CheckDest(ByVal OrgCode As String, ByVal MaterialNo As String, ByVal RTLot As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As DataSet
        CheckDest = New DataSet
        Dim CheckData As DataTable
        Dim myDataColumn As DataColumn
        Dim myDataRow As Data.DataRow
        Dim l_check As Boolean

        CheckData = New Data.DataTable("CheckData")
        myDataColumn = New Data.DataColumn("SubRc", System.Type.GetType("System.Int16"))
        CheckData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ErrMSG", System.Type.GetType("System.String"))
        CheckData.Columns.Add(myDataColumn)
        CheckDest.Tables.Add(CheckData)
        myDataRow = CheckDest.Tables("CheckData").NewRow()

        If DestSubInv <> "" And Not DestSubInv Like "%OS%" And DestLocator <> "" Then             'If destination is not MRB sub inventory then it will check
            'Zhoongshan need to check if same material is sitting in destination
            Dim org_code As String
            Dim ds As New DataSet
            org_code = OracleLoginData.OrgCode

            Try
                Using da As DataAccess = GetDataAccess()
                    Dim Sqlstr As String
                    If org_code <> "353" And org_code <> "354" And org_code <> "355" Then  'FY + LD + ZhongShan
                        'ErrorLogging(MoveType & "CheckDest", OracleLoginData.User, MaterialNo & RTLot & DestSubInv & DestLocator, "I")
                        Sqlstr = String.Format("SELECT TOP (1) CLID, SLOC, StorageBin, MaterialNo, RTLot  FROM T_CLMaster with (nolock) WHERE  StatusCode = '1' and MaterialNo = '{0}' and RTLot <> '{1}' and SLOC = '{2}' and StorageBin = '{3}' and OrgCode = '{4}'", MaterialNo, RTLot, DestSubInv, DestLocator, org_code)
                        ds = da.ExecuteDataSet(Sqlstr, "DestData")
                    Else    'PHl + Vitnam ?
                        Sqlstr = String.Format("SELECT TOP (1) CLID, SLOC, StorageBin, MaterialNo, RTLot  FROM T_CLMaster with (nolock) WHERE  StatusCode = '1' and MaterialNo = '{0}' and SLOC  = '{1}' and StorageBin = '{2}' and OrgCode = '{3}'", MaterialNo, DestSubInv, DestLocator, org_code)
                        ds = da.ExecuteDataSet(Sqlstr, "DestData")
                        If ds.Tables.Count > 0 Then
                            If ds.Tables("DestData").Rows.Count > 0 Then
                                l_check = True
                            Else
                                l_check = False
                                ds.Tables.Clear()
                                Sqlstr = String.Format("SELECT TOP (1) CLID, SLOC, StorageBin, MaterialNo, RTLot  FROM T_CLMaster with (nolock) WHERE  StatusCode = '1' and MaterialNo = '{0}' and SLOC  = '{1}' and OrgCode = '{2}'", MaterialNo, DestSubInv, org_code)
                                ds = da.ExecuteDataSet(Sqlstr, "DestData")
                            End If
                        End If
                    End If
                End Using

                If ds.Tables.Count > 0 Then  'If data is found, 
                    If ds.Tables("DestData").Rows.Count > 0 Then
                        If org_code <> "353" And org_code <> "354" And org_code <> "355" Then  'FY + LD + ZhongShan
                            'ErrorLogging(MoveType & "CheckDest", OracleLoginData.User, ds.Tables("DestData").Rows(0)("CLID"), "I")
                            myDataRow("SubRc") = 2
                            myDataRow("ErrMSG") = "Different LotNo exists in the same locator. Pls putaway to another locator!"
                            CheckDest.Tables("CheckData").Rows.Add(myDataRow)
                        Else 'PHl + Vitnam ?
                            If l_check = False Then
                                myDataRow("SubRc") = 1  'Warning message
                                myDataRow("ErrMSG") = "Material exists in " + ds.Tables("DestData").Rows(0)("SLOC").ToString + " " + ds.Tables("DestData").Rows(0)("StorageBin").ToString + ". [Yes] to continue. [No] to change locator."
                                CheckDest.Tables("CheckData").Rows.Add(myDataRow)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Throw ex
                ErrorLogging(MoveType & "CheckDest", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End If
    End Function

    Public Function Check_Dest(ByVal OrgCode As String, ByVal MaterialNo As String, ByVal RTLot As String, ByVal DestSubInv As String, ByVal DestLocator As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin) As String
        Dim check_lot As String
        If FixNull(DestSubInv) <> "" AndAlso InStr(FixNull(DestSubInv), "OS") < 1 AndAlso InStr(FixNull(DestSubInv), "FS") < 1 AndAlso MoveType <> "MMt-SubCost" Then
            Using da As DataAccess = GetDataAccess()
                Try
                    check_lot = Convert.ToString(da.ExecuteScalar(String.Format("select Value from T_Config with (nolock) where ConfigID = 'MMT001'")))
                    If Not check_lot Is Nothing AndAlso Not check_lot Is DBNull.Value AndAlso check_lot.ToString = "YES" Then
                        Dim Sqlstr As String
                        Dim checkdata As DataSet = New DataSet
                        Sqlstr = String.Format("SELECT TOP (1) CLID, SLOC, StorageBin, MaterialNo, RTLot  FROM T_CLMaster with (nolock) WHERE  StatusCode = '1' and MaterialNo = '{0}' and RTLot <> '{1}' and SLOC = '{2}' and StorageBin = '{3}' and OrgCode = '{4}'", Trim(MaterialNo), Trim(RTLot), FixNull(DestSubInv), FixNull(DestLocator), OrgCode)
                        checkdata = da.ExecuteDataSet(Sqlstr, "DestData")
                        If Not checkdata Is Nothing AndAlso checkdata.Tables.Count > 0 AndAlso checkdata.Tables(0).Rows.Count > 0 Then
                            Check_Dest = "Different LotNo exists in the same locator. Pls transfer to another locator!"
                        End If
                    End If
                Catch ex As Exception
                    Throw ex
                    ErrorLogging(MoveType & "Check_Dest", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try
            End Using
        End If
    End Function

    Public Function Check_Return_Option(ByVal OracleLoginData As ERPLogin) As String
        Dim result As String
        Using da As DataAccess = GetDataAccess()
            Try
                result = Convert.ToString(da.ExecuteScalar(String.Format("select Value from T_Config with (nolock) where ConfigID = 'MMT002'")))
            Catch ex As Exception
                Throw ex
                ErrorLogging("Check_Return_Option", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
        Check_Return_Option = FixNull(result)
    End Function

    Public Function Misc_issue_rcpt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv  '20560
        Dim appl As Integer = OracleLoginData.AppID_Inv  '706
        Dim acc_type As String
        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            If MoveType = "MMt-ProdRtn" Then
                acc_type = "PROD"
            ElseIf MoveType = "MMt-OSPRtn" Then
                acc_type = "OSP"
            End If
            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_wip_pkg.misc_issue_rcpt"
            comm.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_rev", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_num", OracleType.VarChar, 240)
            comm.Parameters.Add("p_quantity", OracleType.Double)
            comm.Parameters.Add("p_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_sub", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_loc", OracleType.VarChar, 240)
            comm.Parameters.Add("p_dest_sub", OracleType.VarChar, 240)
            comm.Parameters.Add("p_dest_loc", OracleType.VarChar, 240)
            comm.Parameters.Add("p_exp_date", OracleType.DateTime)
            comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
            comm.Parameters.Add("p_reference", OracleType.VarChar, 240)
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, acc_type & "0", "I")
            comm.Parameters.Add("p_acc_type", OracleType.VarChar, 240).Value = acc_type
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, acc_type & "1", "I")
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 240)

            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            'comm.Parameters("p_org_code").SourceColumn = "p_org_code"
            comm.Parameters("p_item_num").SourceColumn = "p_item_num"
            comm.Parameters("p_item_rev").SourceColumn = "p_item_rev"
            comm.Parameters("p_lot_num").SourceColumn = "p_lot_num"
            comm.Parameters("p_quantity").SourceColumn = "p_quantity"
            comm.Parameters("p_uom").SourceColumn = "p_uom"
            comm.Parameters("p_source_sub").SourceColumn = "p_source_sub"
            comm.Parameters("p_source_loc").SourceColumn = "p_source_loc"
            comm.Parameters("p_dest_sub").SourceColumn = "p_dest_sub"
            comm.Parameters("p_dest_loc").SourceColumn = "p_dest_loc"
            comm.Parameters("p_exp_date").SourceColumn = "p_exp_date"
            comm.Parameters("p_reason").SourceColumn = "p_reason"
            comm.Parameters("p_reference").SourceColumn = "p_reference"
            'comm.Parameters("p_acc_type").SourceColumn = "p_acc_type"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("component_return_table"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("component_return_table").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType, OracleLoginData.User, DR(i)("o_error_mssg").ToString, "I")
                Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If
        Catch ex As Exception
            p_ds.Tables("component_return_table").Rows(0)("o_success_flag") = "N"
            p_ds.Tables("component_return_table").Rows(0)("o_error_mssg") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function post_push_return(ByVal MoveOracle As DataSet, ByVal CLID As String, ByVal Qty As Decimal, ByVal DestSub As String, ByVal DestLoc As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin, ByVal SlotNo As String) As Result
        Dim Result_Push As New DataSet
        Dim Done_Flag As Boolean

        Try
            Result_Push = component_return(MoveOracle, OracleLoginData, MoveType)
            If Result_Push.Tables("component_return_table").Rows(0)("o_success_flag") = "N" Then
                post_push_return.Result_Flag = "N"
                post_push_return.ErrMsg = Result_Push.Tables("component_return_table").Rows(0)("o_error_mssg")
            Else
                Done_Flag = ChangeCLID(CLID, Qty, DestSub, DestLoc, OracleLoginData, MoveType, SlotNo)
                If Done_Flag = True Then
                    post_push_return.Result_Flag = "Y"
                    post_push_return.ErrMsg = ""
                Else
                    post_push_return.Result_Flag = "N"
                    post_push_return.ErrMsg = "CLID record updating error!"
                End If
            End If
        Catch ex As Exception
            If Result_Push.Tables("component_return_table").Rows(0)("o_success_flag") = "N" Then
                post_push_return.Result_Flag = "N"
                post_push_return.ErrMsg = "Oracle posting error:" & post_push_return.ErrMsg
            ElseIf Result_Push.Tables("component_return_table").Rows(0)("o_success_flag") = "Y" Then
                If Done_Flag = False Then
                    post_push_return.Result_Flag = "N"
                    post_push_return.ErrMsg = "CLID record updating error!"
                Else
                    post_push_return.Result_Flag = "Y"
                    post_push_return.ErrMsg = ""
                End If
            End If
        End Try
    End Function

    Public Function component_return(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '20560  'CAROLD3    54050
        Dim appl As Integer = OracleLoginData.AppID_Inv   '706

        Try

            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_wip_pkg.process_cmpnt_return"
            comm.Parameters.Add("p_dj_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_rev", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_return_quantity", OracleType.Double)
            comm.Parameters.Add("p_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
            comm.Parameters.Add("p_ref", OracleType.VarChar, 240)
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 240)

            comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

            comm.Parameters("p_dj_name").SourceColumn = "p_dj_name"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_item_num").SourceColumn = "p_item_num"
            comm.Parameters("p_item_rev").SourceColumn = "p_item_rev"
            comm.Parameters("p_subinventory").SourceColumn = "p_subinventory"
            comm.Parameters("p_locator").SourceColumn = "p_locator"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_return_quantity").SourceColumn = "p_return_quantity"
            comm.Parameters("p_uom").SourceColumn = "p_uom"
            comm.Parameters("p_reason").SourceColumn = "p_reason"
            comm.Parameters("p_ref").SourceColumn = "p_ref"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("component_return_table"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("component_return_table").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType, OracleLoginData.User, DR(i)("o_error_mssg").ToString, "I")
                Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If
        Catch ex As Exception
            p_ds.Tables("component_return_table").Rows(0)("o_success_flag") = "N"
            p_ds.Tables("component_return_table").Rows(0)("o_error_mssg") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function del_mtl_tran_inte(ByVal p_type As String, ByVal p_OracleLoginData As ERPLogin, ByVal p_transaction_header_id As Integer, ByVal p_timeout As Integer) As String
        'del_rcv_inte(p_group_id number,p_user_id number,o_succ_flag out varchar2,o_return_message out varchar2)
        Using da As DataAccess = GetDataAccess()
            Dim Oda As OracleCommand = da.OraCommand()

            Try
                Dim aa As Integer
                Dim MsgFlag As String
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.del_mtl_tran_inte"
                Oda.Parameters.Add("p_tran_head_id", OracleType.Int32).Value = p_transaction_header_id
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = p_OracleLoginData.UserID
                'Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = GetOrgCode()   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("o_succ_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_return_message", OracleType.VarChar, 240)
                Oda.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_return_message").Direction = ParameterDirection.Output

                Oda.Connection.Open()
                aa = CInt(Oda.ExecuteNonQuery())
                MsgFlag = Oda.Parameters("o_succ_flag").Value
                Return DirectCast(MsgFlag, String)
                Oda.Connection.Close()

            Catch oe As OracleException
                ErrorLogging(p_type, p_OracleLoginData.User, "TransactionID: " & p_transaction_header_id & ", " & oe.Message & oe.Source, "E")
                Throw oe
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function transfer_mat_submit(ByVal p_type As String, ByVal p_OracleLoginData As ERPLogin, ByVal p_transaction_header_id As Integer, ByVal p_timeout As Integer) As DataSet
        Dim p_ds As New DataSet()
        p_ds.Tables.Add("MSG")
        Dim oda As New OracleDataAdapter()
        Dim comm_submit As New OracleCommand()
        Dim oc As New OracleConnection("Data Source=CAROLD3;user id=APPS;password=apps")

        Try

            comm_submit.Connection = oc
            oda.SelectCommand = comm_submit
            oda.SelectCommand.CommandType = CommandType.StoredProcedure

            oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.Call_Process_OnLine"
            ' oda.SelectCommand.Parameters.Add("o_error_cursor", OracleType.Cursor).Direction = ParameterDirection.Output

            oda.SelectCommand.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
            oda.SelectCommand.Parameters("p_transaction_header_id").Value = p_transaction_header_id
            oda.SelectCommand.Parameters.Add(New OracleParameter("p_timeout", OracleType.Int32))
            oda.SelectCommand.Parameters("p_timeout").Value = p_timeout

            oda.SelectCommand.Parameters.Add(New OracleParameter("p_error_code", OracleType.VarChar, 240))
            oda.SelectCommand.Parameters.Add(New OracleParameter("p_error_explanation", OracleType.VarChar, 240))
            oda.SelectCommand.Parameters.Add(New OracleParameter("p_process_online_yn", OracleType.VarChar, 240))
            oda.SelectCommand.Parameters.Add(New OracleParameter("x_return_status", OracleType.VarChar, 240))
            oda.SelectCommand.Parameters.Add(New OracleParameter("x_msg_count", OracleType.Int32))
            oda.SelectCommand.Parameters.Add(New OracleParameter("x_msg_data", OracleType.VarChar, 240))
            oda.SelectCommand.Parameters.Add(New OracleParameter("x_trans_count", OracleType.Int32))
            oda.SelectCommand.Parameters.Add(New OracleParameter("p_tran_result", OracleType.Int32))
            oda.SelectCommand.Parameters.Add(New OracleParameter("p_error_message", OracleType.Cursor))

            oda.SelectCommand.Parameters("p_error_code").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("p_error_explanation").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("p_process_online_yn").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("x_return_status").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("x_msg_count").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("x_msg_data").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("x_trans_count").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("p_tran_result").Direction = ParameterDirection.Output
            oda.SelectCommand.Parameters("p_error_message").Direction = ParameterDirection.Output

            oda.SelectCommand.Connection.Open()
            oda.Fill(p_ds, "MSG")
            oda.SelectCommand.Connection.Close()
            ' oda.Dispose()

        Catch oe As OracleException
            ErrorLogging(p_type, p_OracleLoginData.User, oe.Message & oe.Source, "E")
            Throw oe
        Finally
            If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
        End Try
        Return p_ds
    End Function

    Public Function account_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv  '53485   'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv   '401

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_receipt"      '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
            comm.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountreceipt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountreceipt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function account_alias_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485   'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv   '401
        'Dim Account_Alias As String

        ''Account_Alias = ""
        'If MiscType = "E&O WRITE OFF" Then
        '    Account_Alias = "E&O WRITE OFF"
        'ElseIf MiscType = "EXCESS REIPT FR PROD" Then
        '    Account_Alias = "EXCESS RECPT FM PROD"
        'ElseIf MiscType = "MISC RECEIPT FOR RM" Or MiscType = "MISC RECEIPT FOR SA" Or MiscType = "MISC RECEIPT FOR FG" Then
        '    'ErrorLogging(MoveType, OracleLoginData.User, "test01")
        '    Account_Alias = "RMA RECEIPT (PSG)"   '"EXCESS RECPT FM PROD"    '"EXCESS REIPT FR PROD"
        'End If

        'If MoveType = "MMt-ProdRtn" Then
        '    Account_Alias = "EXCESS RECPT FM PROD"
        'ElseIf MoveType = "MMt-OSPRtn" Then
        '    Account_Alias = "EXCESS RECPT FM OSP"
        'End If

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            'Dim j As Integer
            'For j = 0 To p_ds.Tables(0).Rows.Count - 1
            '    p_ds.Tables(0).Rows(j)("p_transaction_source_name") = Account_Alias
            'Next

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_receipt"      '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            'comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240).Value = Account_Alias
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)

            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
            comm.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountreceipt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountreceipt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function account_alias_batch_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485   'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv   '401
        'Dim Account_Alias As String

        ''Account_Alias = ""
        'If MiscType = "E&O WRITE OFF" Then
        '    Account_Alias = "E&O WRITE OFF"
        'ElseIf MiscType = "EXCESS REIPT FR PROD" Then
        '    Account_Alias = "EXCESS RECPT FM PROD"
        'ElseIf MiscType = "MISC RECEIPT FOR RM" Or MiscType = "MISC RECEIPT FOR SA" Or MiscType = "MISC RECEIPT FOR FG" Then
        '    'ErrorLogging(MoveType, OracleLoginData.User, "test01", "I")
        '    Account_Alias = "RMA RECEIPT (PSG)"   '"EXCESS RECPT FM PROD"    '"EXCESS REIPT FR PROD"
        'End If

        'If MoveType = "MMt-ProdRtn" Then
        '    Account_Alias = "EXCESS RECPT FM PROD"
        'ElseIf MoveType = "MMt-OSPRtn" Then
        '    Account_Alias = "EXCESS RECPT FM OSP"
        'End If

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            Dim j As Integer
            'For j = 0 To p_ds.Tables(0).Rows.Count - 1
            '    p_ds.Tables(0).Rows(j)("p_transaction_source_name") = Account_Alias
            'Next
            For j = 0 To p_ds.Tables("accountreceipt_table").Rows.Count - 1
                If p_ds.Tables("accountreceipt_table").Rows(j).RowState = DataRowState.Unchanged Then
                    p_ds.Tables("accountreceipt_table").Rows(j).SetAdded()
                End If
            Next

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_batch_receipt"      '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            'comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240).Value = Account_Alias
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
            comm.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountreceipt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            Dim result_flag As String
            result_flag = ""

            DR = p_ds.Tables("accountreceipt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                'comm.Transaction.Commit()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds
                result_flag = submit_inv_cvs("account_alias_batch_receipt", TransactionID, OracleLoginData)

                If result_flag = "Y" Then

                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag, "I")
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_inv_cvs("account_alias_batch_receipt", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next
                'comm.Transaction.Rollback()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds

                result_flag = del_inv_cvs("account_alias_batch_receipt", TransactionID, OracleLoginData)
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function account_alias_batch_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485   'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv   '401

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            Dim j As Integer
            'For j = 0 To p_ds.Tables(0).Rows.Count - 1
            '    p_ds.Tables(0).Rows(j)("p_transaction_source_name") = Account_Alias
            'Next
            For j = 0 To p_ds.Tables("accountissue_table").Rows.Count - 1
                If p_ds.Tables("accountissue_table").Rows(j).RowState = DataRowState.Unchanged Then
                    p_ds.Tables("accountissue_table").Rows(j).SetAdded()
                End If
            Next

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_batch_issue"      '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            'comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
            comm.Parameters("p_locator_source").SourceColumn = "p_locator_source"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            'comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountissue_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            Dim result_flag As String
            result_flag = ""

            DR = p_ds.Tables("accountissue_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                'comm.Transaction.Commit()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds
                result_flag = submit_inv_cvs("account_alias_batch_issue", TransactionID, OracleLoginData)

                If result_flag = "Y" Then

                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag, "I")
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_inv_cvs("account_alias_batch_issue", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next
                'comm.Transaction.Rollback()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds

                result_flag = del_inv_cvs("account_alias_batch_issue", TransactionID, OracleLoginData)
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function submit_inv_cvs(ByVal MoveType As String, ByVal TransactionID As Double, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As OracleString
            Dim bb As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand
            Dim transaction As OracleTransaction
            Dim resp As Integer = OracleLoginData.RespID_Inv
            Dim appl As Integer = OracleLoginData.AppID_Inv

            Try
                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If
                transaction = Oda.Connection.BeginTransaction(IsolationLevel.ReadCommitted)
                Oda.Transaction = transaction

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = OracleLoginData.RespID_Inv  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = OracleLoginData.AppID_Inv
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.mtl_process_online"
                Oda.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                'Oda.Parameters.Add("p_timout", OracleType.Int32).Value = 1800000
                'Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = GetOrgCode()   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_error_mssg", OracleType.VarChar, 10000)
                Oda.Parameters("o_success_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_error_mssg").Direction = ParameterDirection.Output

                bb = CInt(Oda.ExecuteOracleNonQuery(aa))
                MsgFlag = Oda.Parameters("o_success_flag").Value
                If MsgFlag <> "Y" Then
                    MsgFlag = Oda.Parameters("o_error_mssg").Value
                    transaction.Rollback()
                Else
                    transaction.Commit()
                End If
                Oda.Connection.Close()
                Return DirectCast(MsgFlag, String)
            Catch oe As OracleException
                ErrorLogging(MoveType, OracleLoginData.User, "TransactionID: " & TransactionID & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function del_inv_cvs(ByVal MoveType As String, ByVal TransactionID As Double, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()

            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.del_mtl_tran_inte"
                Oda.Parameters.Add("p_tran_head_id", OracleType.Double).Value = TransactionID
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = OracleLoginData.UserID
                'Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = GetOrgCode()   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("o_succ_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_return_message", OracleType.VarChar, 240)
                Oda.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_return_message").Direction = ParameterDirection.Output

                Oda.Connection.Open()
                aa = CInt(Oda.ExecuteNonQuery())
                MsgFlag = Oda.Parameters("o_return_message").Value
                Return DirectCast(MsgFlag, String)
                Oda.Connection.Close()

            Catch oe As OracleException
                ErrorLogging(MoveType, OracleLoginData.User, "TransactionID: " & TransactionID & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function account_alias_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Dim Account_Alias As String

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            If MoveType = "MMt-ProdRtn" Then
                Account_Alias = "EXCESS ISSUE TO PROD"
            ElseIf MoveType = "MMt-OSPRtn" Then
                Account_Alias = "EXCESS ISSUE TO OSP"
            End If


            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_issue"   '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240).Value = Account_Alias
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            'comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
            comm.Parameters("p_locator_source").SourceColumn = "p_locator_source"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountissue_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountissue_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            'comm.Transaction.Rollback()
            comm.Transaction.Commit()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function post_pull_return(ByVal MoveOracle As DataSet, ByVal TransactionID As Long, ByVal MiscType As String, ByVal CLID As String, ByVal Qty As Decimal, ByVal DestSub As String, ByVal DestLoc As String, ByVal MoveType As String, ByVal OracleLoginData As ERPLogin, ByVal SlotNo As String) As Result
        Dim Result_Pull As New DataSet
        Dim Done_Flag As Boolean

        Try
            Result_Pull = account_issue_rcpt(MoveOracle, OracleLoginData, MoveType, TransactionID, MiscType)
            If Result_Pull.Tables("issuercpt_table").Rows(0)("o_return_status") = "N" Then
                post_pull_return.Result_Flag = "N"
                post_pull_return.ErrMsg = Result_Pull.Tables("issuercpt_table").Rows(0)("o_return_message")
            Else
                Done_Flag = ChangeCLID(CLID, Qty, DestSub, DestLoc, OracleLoginData, MoveType, SlotNo)
                If Done_Flag = True Then
                    post_pull_return.Result_Flag = "Y"
                    post_pull_return.ErrMsg = ""
                Else
                    post_pull_return.Result_Flag = "N"
                    post_pull_return.ErrMsg = "CLID record updating error!"
                End If
            End If
        Catch ex As Exception
            If Result_Pull.Tables("issuercpt_table").Rows(0)("o_return_status") = "N" Then
                post_pull_return.Result_Flag = "N"
                post_pull_return.ErrMsg = "Oracle posting error:" & post_pull_return.ErrMsg
            ElseIf Result_Pull.Tables("issuercpt_table").Rows(0)("o_return_status") = "Y" Then
                If Done_Flag = False Then
                    post_pull_return.Result_Flag = "N"
                    post_pull_return.ErrMsg = "CLID record updating error!"
                Else
                    post_pull_return.Result_Flag = "Y"
                    post_pull_return.ErrMsg = ""
                End If
            End If
        End Try
    End Function

    Public Function account_issue_rcpt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Dim Account_Alias_Issue As String
        Dim Account_Alias_Rcpt As String

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            'If MoveType = "MMt-ProdRtn" Then
            '    Account_Alias_Issue = "EXCESS ISSUE TO PROD"
            '    Account_Alias_Rcpt = "EXCESS RECPT FM PROD"
            'ElseIf MoveType = "MMt-OSPRtn" Then
            '    Account_Alias_Issue = "EXCESS ISSUE TO OSP"
            '    Account_Alias_Rcpt = "EXCESS RECPT FM OSP"
            'End If
            If MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Then
                Account_Alias_Issue = "MAT ISSUE (PRDN RET)"
                Account_Alias_Rcpt = "MAT RCPT (PRDN RET)"
            ElseIf MoveType = "MMt-OSPRtn" Then
                Account_Alias_Issue = "MAT ISSUE (OSP RET)"
                Account_Alias_Rcpt = "MAT RCPT (OSP RET)"
            End If
            TransactionID = CLng(GetNextHeaderID(OracleLoginData))
            p_ds.Tables("issuercpt_table").Rows(0)("p_transaction_header_id") = TransactionID

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_issue_rcpt"   '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_issue_tran_source_name", OracleType.VarChar, 240).Value = Account_Alias_Issue
            comm.Parameters.Add("p_rcpt_tran_source_name", OracleType.VarChar, 240).Value = Account_Alias_Rcpt
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
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            'comm.Parameters("p_issue_tran_source_name").SourceColumn = "p_issue_tran__source_name"
            'comm.Parameters("p_rcpt_tran_source_name").SourceColumn = "p_rcpt_tran_source_name"
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
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("issuercpt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("issuercpt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next
                comm.Transaction.Rollback()
                comm.Connection.Close()
                del_inv_cvs("account_issue_rcpt", TransactionID, OracleLoginData)
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("issuercpt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("issuercpt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function account_issue_rcpt_v2(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Dim Account_Alias_Issue As String
        Dim Account_Alias_Rcpt As String

        Try
            'If MoveType = "MMt-ProdRtn" Then
            '    Account_Alias_Issue = "EXCESS ISSUE TO PROD"
            '    Account_Alias_Rcpt = "EXCESS RECPT FM PROD"
            'ElseIf MoveType = "MMt-OSPRtn" Then
            '    Account_Alias_Issue = "EXCESS ISSUE TO OSP"
            '    Account_Alias_Rcpt = "EXCESS RECPT FM OSP"
            'End If
            If MoveType = "MMt-ProdRtn" Or MoveType = "MMt-FSRtn" Then
                Account_Alias_Issue = "MAT ISSUE (PRDN RET)"
                Account_Alias_Rcpt = "MAT RCPT (PRDN RET)"
            ElseIf MoveType = "MMt-OSPRtn" Then
                Account_Alias_Issue = "MAT ISSUE (OSP RET)"
                Account_Alias_Rcpt = "MAT RCPT (OSP RET)"
            End If
            TransactionID = CLng(GetNextHeaderID(OracleLoginData))
            p_ds.Tables("issuercpt_table").Rows(0)("p_transaction_header_id") = TransactionID

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_issue_rcpt_v2"   '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_issue_tran_source_name", OracleType.VarChar, 240).Value = Account_Alias_Issue
            comm.Parameters.Add("p_rcpt_tran_source_name", OracleType.VarChar, 240).Value = Account_Alias_Rcpt
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
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            'comm.Parameters("p_issue_tran_source_name").SourceColumn = "p_issue_tran__source_name"
            'comm.Parameters("p_rcpt_tran_source_name").SourceColumn = "p_rcpt_tran_source_name"
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
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("issuercpt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            Dim result_flag As String
            result_flag = ""

            DR = p_ds.Tables("issuercpt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                result_flag = submit_inv_cvs("account_alias_issue_receipt", TransactionID, OracleLoginData)

                If result_flag = "Y" Then

                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag)
                    p_ds.Tables("issuercpt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("issuercpt_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_inv_cvs("account_alias_issue_receipt_v2", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString)
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString)
                'Next
                result_flag = del_inv_cvs("account_alias_issue_receipt_v2", TransactionID, OracleLoginData)
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("issuercpt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("issuercpt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source)
            comm.Transaction.Rollback()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function account_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_issue"   '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
            comm.Parameters("p_locator_source").SourceColumn = "p_locator_source"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountissue_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountissue_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            'comm.Transaction.Rollback()
            comm.Transaction.Commit()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function post_misc_rcpt(ByVal MoveOracle As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal LabelInfo As LabelData, ByVal Pkg As Integer, ByVal Printer As String) As ConversionResult
        Try
            Dim DR() As DataRow = Nothing
            Dim dtRow As DataRow
            Dim msg As String
            Dim i As Integer
            Dim post_result As DataSet

            post_misc_rcpt.OracleFlag = "N"
            post_misc_rcpt.CLIDFlag = "N"
            post_misc_rcpt.PrintFlag = "N"

            post_result = account_alias_receipt(MoveOracle, OracleLoginData, "", TransactionID, "")
            DR = post_result.Tables("accountreceipt_table").Select("o_return_status = 'N'")
            If DR.Length = 0 Then
                post_misc_rcpt.OracleFlag = "Y"
                post_misc_rcpt = CLIDforMiscRcpt(OracleLoginData, MoveType, LabelInfo, Pkg, Printer)
            Else
                For Each dtRow In DR
                    msg = msg & dtRow("o_return_message").ToString
                Next
                post_misc_rcpt.ErrorMsg = msg
                'post_misc_rcpt.OracleFlag = "N"
                'post_misc_rcpt.CLIDFlag = "N"
                'post_misc_rcpt.PrintFlag = "N"
            End If
        Catch ex As Exception
            'post_misc_rcpt.OracleFlag = "N"
            'post_misc_rcpt.CLIDFlag = "N"
            post_misc_rcpt.ErrorMsg = ex.Message.ToString
            ErrorLogging("post_misc_rcpt", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
        Return post_misc_rcpt
    End Function

    Public Function CLIDforMiscRcpt(ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal LabelInfo As LabelData, ByVal Pkg As Integer, ByVal Printer As String) As ConversionResult
        Dim i, k As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim strCMD As String
        Dim NextCLID As String
        Dim Blank As String
        Dim T_Flag As String
        Dim ExpDate As String
        Dim V_Null As DBNull
        Dim PrintFlag As Boolean
        Dim TraceType As String
        Dim LastTransaction As String
        Dim myCommand As SqlClient.SqlCommand

        Dim CLIDs As New DataSet                     'Set CLIDs DataSet to save table
        Dim CLIDsTable As DataTable                  'Set CLIDList Table to save CLIDs
        Dim myDataColumn As DataColumn
        Dim myDataRow As Data.DataRow

        CLIDs = New DataSet
        CLIDsTable = New Data.DataTable("CLIDList")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CLIDsTable.Columns.Add(myDataColumn)
        CLIDs.Tables.Add(CLIDsTable)
        Try
            CLIDforMiscRcpt.OracleFlag = "Y"
            CLIDforMiscRcpt.CLIDFlag = "N"
            CLIDforMiscRcpt.PrintFlag = "N"
            Blank = ""
            T_Flag = "True"
            V_Null = DBNull.Value
            TraceType = "PT"
            LastTransaction = "MiscRcpt_WithoutCLID"
            If Not LabelInfo.ExpDate Is Nothing AndAlso FixNull(LabelInfo.ExpDate) <> "" Then
                ExpDate = Replace(LabelInfo.ExpDate, "-", "/")
            Else
                ExpDate = ""
            End If
            myConn.Open()
            For i = 0 To Pkg - 1
                If MoveType = "MiscRcpt_RM" Then
                    NextCLID = GetNextCLID(OracleLoginData)
                Else
                    NextCLID = GetNextProdID(OracleLoginData)
                End If
                'If Not LabelInfo.RefCLID Is Nothing AndAlso LabelInfo.RefCLID.Tables.Count > 0 AndAlso LabelInfo.RefCLID.Tables(0).Rows.Count > 0 Then
                If FixNull(LabelInfo.RefCLID) <> "" Then
                    Dim sqldatenull As DateTime
                    Dim smalltime As Date


                    strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,ProcessID,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,DeliveryType,VendorName,VendorPN,SLOC,StoragePosition,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,NextReviewDate,AddlData,Stemp,MSL,LastTransaction,ItemText) select '{0}','{1}', '{11}', MaterialNo, MaterialRevision, 0, UOM, '{2}', BaseUOM, DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,'{3}',getdate(),'{12}',RecDate,Printed,ExpDate,ProdDate,ProcessID,RoHS,PurOrdNo,PurOrdItem,'{10}',BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,DeliveryType,VendorName,VendorPN,'{4}',StoragePosition,'{5}','{6}',Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,NextReviewDate,AddlData,Stemp,MSL,'{7}','{9}' from T_Clmaster as T_Clmaster1 where CLID = '{8}'", NextCLID, OracleLoginData.OrgCode, LabelInfo.Qty, OracleLoginData.User, LabelInfo.SubInv, LabelInfo.StorageType, LabelInfo.Locator, "MiscRcpt_WithoutCLID", LabelInfo.RefCLID, LabelInfo.ItemText, LabelInfo.InvoiceNo, LabelInfo.StatusCode, LabelInfo.VendorID)

                    'If ExpDate <> "" AndAlso LabelInfo.RTLot <> "" Then
                    '    strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction) values ('{0}', '{1}', 1, '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{23}','{8}','{23}','{8}','{9}',getdate(),'{24}',getdate(),'{20}','{10}',getdate(),'{11}','{24}','{24}','{24}','{24}','{24}','{24}','{24}','{24}','{12}','{24}','{24}','{13}','{14}','{15}','{21}','{8}','{23}','{16}','{17}','{23}','{23}','{23}','{19}','{22}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, LabelInfo.DCode, LabelInfo.LotNo, LabelInfo.RTNo, OracleLoginData.User, ExpDate, LabelInfo.RoHS, LabelInfo.Description, LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, LabelInfo.MFR, LabelInfo.MPN, LabelInfo.Stemp, LabelInfo.MSL, "True", "PT", "MiscRcpt_WithoutCLID", V_Null, Blank)
                    'ElseIf ExpDate <> "" AndAlso LabelInfo.RTLot = "" Then
                    '    strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction) values ('{0}', '{1}', 1, '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{23}','{8}','{23}','{8}','{9}',getdate(),'{24}',getdate(),'{20}','{10}',getdate(),'{11}','{24}','{24}','{24}','{24}','{24}','{24}','{24}','{24}','{12}','{24}','{24}','{13}','{14}','{15}','{21}','{8}','{23}','{16}','{17}','{23}','{23}','{23}','{19}','{22}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, LabelInfo.DCode, LabelInfo.LotNo, Blank, OracleLoginData.User, ExpDate, LabelInfo.RoHS, LabelInfo.Description, LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, LabelInfo.MFR, LabelInfo.MPN, LabelInfo.Stemp, LabelInfo.MSL, "True", "PT", "MiscRcpt_WithoutCLID", V_Null, Blank)
                    'ElseIf ExpDate = "" AndAlso LabelInfo.RTLot <> "" Then
                    '    strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction) values ('{0}', '{1}', 1, '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{22}','{8}','{22}','{8}','{9}',getdate(),'{23}',getdate(),'{19}',getdate(),'{10}','{23}','{23}','{23}','{23}','{23}','{23}','{23}','{23}','{11}','{23}','{23}','{12}','{13}','{14}','{20}','{8}','{22}','{15}','{16}','{22}','{22}','{22}','{18}','{21}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, LabelInfo.DCode, LabelInfo.LotNo, LabelInfo.RTNo, OracleLoginData.User, LabelInfo.RoHS, LabelInfo.Description, LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, LabelInfo.MFR, LabelInfo.MPN, LabelInfo.Stemp, LabelInfo.MSL, "True", "PT", "MiscRcpt_WithoutCLID", DBNull.Value, Blank)
                    'ElseIf ExpDate = "" AndAlso LabelInfo.RTLot = "" Then
                    '    strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction) values ('{0}', '{1}', 1, '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{22}','{8}','{22}','{23}','{9}',getdate(),'{23}',getdate(),'{19}',getdate(),'{10}','{23}','{23}','{23}','{23}','{23}','{23}','{23}','{23}','{11}','{23}','{23}','{12}','{13}','{14}','{20}','{8}','{22}','{15}','{16}','{22}','{22}','{22}','{18}','{21}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, LabelInfo.DCode, LabelInfo.LotNo, Blank, OracleLoginData.User, LabelInfo.RoHS, LabelInfo.Description, LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, LabelInfo.MFR, LabelInfo.MPN, LabelInfo.Stemp, LabelInfo.MSL, "True", "PT", "MiscRcpt_WithoutCLID", DBNull.Value, Blank)
                    'End If
                Else
                    If ExpDate <> "" AndAlso LabelInfo.RTLot <> "" Then
                        strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction,ItemText) values ('{0}', '{1}', '{29}', '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{23}','{8}','{23}','{8}','{9}',getdate(),'{30}',getdate(),'{20}','{10}',getdate(),'{11}','{24}','{24}','{28}','{24}','{24}','{24}','{24}','{25}','{12}','{24}','{24}','{13}','{26}','{14}','{15}','{21}','{8}','{23}','{16}','{17}','{23}','{23}','{23}','{19}','{22}','{27}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, SQLString(LabelInfo.DCode), SQLString(LabelInfo.LotNo), LabelInfo.RTNo, OracleLoginData.User, ExpDate, SQLString(LabelInfo.RoHS), SQLString(LabelInfo.Description), LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, SQLString(LabelInfo.MFR), SQLString(LabelInfo.MPN), SQLString(LabelInfo.Stemp), SQLString(LabelInfo.MSL), "True", "PT", "MiscRcpt_WithoutCLID", V_Null, Blank, "QP", LabelInfo.StorageType, LabelInfo.ItemText, LabelInfo.InvoiceNo, LabelInfo.StatusCode, LabelInfo.VendorID)
                    ElseIf ExpDate <> "" AndAlso LabelInfo.RTLot = "" Then
                        strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction,ItemText) values ('{0}', '{1}', '{29}', '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{23}','{8}','{23}','{8}','{9}',getdate(),'{30}',getdate(),'{20}','{10}',getdate(),'{11}','{24}','{24}','{28}','{24}','{24}','{24}','{24}','{25}','{12}','{24}','{24}','{13}','{26}','{14}','{15}','{21}','{8}','{23}','{16}','{17}','{23}','{23}','{23}','{19}','{22}','{27}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, SQLString(LabelInfo.DCode), SQLString(LabelInfo.LotNo), Blank, OracleLoginData.User, ExpDate, SQLString(LabelInfo.RoHS), SQLString(LabelInfo.Description), LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, SQLString(LabelInfo.MFR), SQLString(LabelInfo.MPN), SQLString(LabelInfo.Stemp), SQLString(LabelInfo.MSL), "True", "PT", "MiscRcpt_WithoutCLID", V_Null, Blank, "QP", LabelInfo.StorageType, LabelInfo.ItemText, LabelInfo.InvoiceNo, LabelInfo.StatusCode, LabelInfo.VendorID)
                    ElseIf ExpDate = "" AndAlso LabelInfo.RTLot <> "" Then
                        strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction,ItemText) values ('{0}', '{1}', '{28}', '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{22}','{8}','{22}','{8}','{9}',getdate(),'{29}',getdate(),'{19}',getdate(),'{10}','{23}','{23}','{27}','{23}','{23}','{23}','{23}','{24}','{11}','{23}','{23}','{12}','{25}','{13}','{14}','{20}','{8}','{22}','{15}','{16}','{22}','{22}','{22}','{18}','{21}','{26}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, SQLString(LabelInfo.DCode), SQLString(LabelInfo.LotNo), LabelInfo.RTNo, OracleLoginData.User, SQLString(LabelInfo.RoHS), SQLString(LabelInfo.Description), LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, SQLString(LabelInfo.MFR), SQLString(LabelInfo.MPN), SQLString(LabelInfo.Stemp), SQLString(LabelInfo.MSL), "True", "PT", "MiscRcpt_WithoutCLID", DBNull.Value, Blank, "QP", LabelInfo.StorageType, LabelInfo.ItemText, LabelInfo.InvoiceNo, LabelInfo.StatusCode, LabelInfo.VendorID)
                    ElseIf ExpDate = "" AndAlso LabelInfo.RTLot = "" Then
                        strCMD = String.Format("Insert into T_CLMaster (CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode,StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction,ItemText) values ('{0}', '{1}', '{28}', '{2}','{3}','{4}','{5}','{4}','{5}','{6}','{7}','{22}','{8}','{22}','{23}','{9}',getdate(),'{29}',getdate(),'{19}',getdate(),'{10}','{23}','{23}','{27}','{23}','{23}','{23}','{23}','{24}','{11}','{23}','{23}','{12}','{25}','{13}','{14}','{20}','{8}','{22}','{15}','{16}','{22}','{22}','{22}','{18}','{21}','{26}')", NextCLID, OracleLoginData.OrgCode, LabelInfo.Material, LabelInfo.MatlRev, LabelInfo.Qty, LabelInfo.UOM, SQLString(LabelInfo.DCode), SQLString(LabelInfo.LotNo), Blank, OracleLoginData.User, SQLString(LabelInfo.RoHS), SQLString(LabelInfo.Description), LabelInfo.SubInv, LabelInfo.Locator, OracleLoginData.User, SQLString(LabelInfo.MFR), SQLString(LabelInfo.MPN), SQLString(LabelInfo.Stemp), SQLString(LabelInfo.MSL), "True", "PT", "MiscRcpt_WithoutCLID", DBNull.Value, Blank, "QP", LabelInfo.StorageType, LabelInfo.ItemText, LabelInfo.InvoiceNo, LabelInfo.StatusCode, LabelInfo.VendorID)
                    End If
                End If
                CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand.ExecuteNonQuery()
                'myConn.Close()

                myDataRow = CLIDsTable.NewRow()
                myDataRow("CLID") = NextCLID
                CLIDsTable.Rows.Add(myDataRow)
            Next
            myConn.Close()
            CLIDforMiscRcpt.CLIDs = New DataSet
            CLIDforMiscRcpt.CLIDs = CLIDs
            CLIDforMiscRcpt.CLIDFlag = "Y"

            PrintFlag = PrintCLIDs(CLIDs, Printer)
            If PrintFlag = True Then
                CLIDforMiscRcpt.PrintFlag = "Y"
                CLIDforMiscRcpt.ErrorMsg = ""
            Else
                CLIDforMiscRcpt.PrintFlag = "N"
                CLIDforMiscRcpt.ErrorMsg = "Printing with error"
            End If
        Catch ex As Exception
            'CLIDforMiscRcpt.OracleFlag = "Y"
            'CLIDforMiscRcpt.CLIDFlag = "N"
            'CLIDforMiscRcpt.PrintFlag = "N"
            CLIDforMiscRcpt.ErrorMsg = ex.Message.ToString
            ErrorLogging("CLIDforMiscRcpt", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then
                myConn.Close()
            End If
        End Try
        Return CLIDforMiscRcpt
    End Function

    Public Function LabelForMiscReceipt(ByVal LoginData As ERPLogin, ByVal LabelPrinter As String, ByVal PrintLabels As Boolean, ByVal LabelInfo As LabelData) As ConversionResult

        Dim ExpDate As String

        ' Get Traceability Level and SafeCode according to Material Group              'Add by Yudy 11/28/2008
        Dim SafeCode As String
        'Dim TraceLevel As MGTraceData
        LabelForMiscReceipt.TraceLevel = "PT"    ' "NT" not needed   "CT" must needed  "PT" optional
        SafeCode = ""
        'If Not GRLabel.Matkl Is Nothing Then
        '    TraceLevel = GetMGTraceData(GRLabel.Matkl)                                 'Yudy 11/28/2008
        '    LabelConversion.TraceLevel = TraceLevel.TraceabilityLevel
        '    If TraceLevel.AddlData = "" Then
        '        SafeCode = ""
        '    ElseIf TraceLevel.AddlData = "S" Then
        '        SafeCode = "Safety"
        '    End If
        'End If


        Dim CLIDs As New DataSet                     'Set CLIDs DataSet to save table
        Dim CLIDsTable As DataTable                  'Set CLIDList Table to save CLIDs
        Dim myDataColumn As DataColumn
        Dim myDataRow As Data.DataRow

        CLIDs = New DataSet
        CLIDsTable = New Data.DataTable("CLIDList")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CLIDsTable.Columns.Add(myDataColumn)
        CLIDs.Tables.Add(CLIDsTable)

        Dim myCommand As SqlClient.SqlCommand
        Dim NewCLIDCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim NextCLID As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@RTLot,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ExpDate,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
        NewCLIDCommand.Parameters.Add("@NewCLID", SqlDbType.VarChar, 20, "CLID")
        NewCLIDCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20, "OrgCode")
        NewCLIDCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 30, "MaterialNo")
        NewCLIDCommand.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 20, "MaterialRevision")
        NewCLIDCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        NewCLIDCommand.Parameters.Add("@UOM", SqlDbType.VarChar, 10, "UOM")
        NewCLIDCommand.Parameters.Add("@QtyBaseUOM", SqlDbType.Decimal, 13, "QtyBaseUOM")
        NewCLIDCommand.Parameters.Add("@BaseUOM", SqlDbType.VarChar, 10, "BaseUOM")
        NewCLIDCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 20, "DateCode")
        NewCLIDCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 20, "LotNo")
        NewCLIDCommand.Parameters.Add("@COO", SqlDbType.VarChar, 20, "COO")
        NewCLIDCommand.Parameters.Add("@RecDocNo", SqlDbType.VarChar, 20, "RecDocNo")
        NewCLIDCommand.Parameters.Add("@RecDocItem", SqlDbType.VarChar, 10, "RecDocItem")
        NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
        NewCLIDCommand.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
        NewCLIDCommand.Parameters.Add("@Printed", SqlDbType.VarChar, 100, "Printed")
        NewCLIDCommand.Parameters.Add("@VendorID", SqlDbType.VarChar, 10, "VendorID")
        NewCLIDCommand.Parameters.Add("@RecDate", SqlDbType.SmallDateTime, 4, "RecDate")
        NewCLIDCommand.Parameters.Add("@RoHS", SqlDbType.VarChar, 10, "RoHS")
        NewCLIDCommand.Parameters.Add("@PurOrdNo", SqlDbType.VarChar, 20, "PurOrdNo")
        NewCLIDCommand.Parameters.Add("@PurOrdItem", SqlDbType.VarChar, 10, "PurOrdItem")
        NewCLIDCommand.Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 25, "InvoiceNo")
        NewCLIDCommand.Parameters.Add("@BillofLading", SqlDbType.VarChar, 25, "BillofLading")
        NewCLIDCommand.Parameters.Add("@DN", SqlDbType.VarChar, 25, "DN")
        NewCLIDCommand.Parameters.Add("@HeaderText", SqlDbType.VarChar, 25, "HeaderText")
        NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
        NewCLIDCommand.Parameters.Add("@ProdDate", SqlDbType.SmallDateTime, 4, "ProdDate")
        NewCLIDCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 4, "ReasonCode")
        NewCLIDCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 10, "StockType")
        NewCLIDCommand.Parameters.Add("@MaterialDesc", SqlDbType.VarChar, 50, "MaterialDesc")
        NewCLIDCommand.Parameters.Add("@VendorName", SqlDbType.VarChar, 50, "VendorName")
        NewCLIDCommand.Parameters.Add("@VendorPN", SqlDbType.VarChar, 50, "VendorPN")
        NewCLIDCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 20, "SLOC")
        NewCLIDCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
        NewCLIDCommand.Parameters.Add("@Operator", SqlDbType.VarChar, 10, "Operator")
        NewCLIDCommand.Parameters.Add("@IsTraceable", SqlDbType.VarChar, 10, "IsTraceable")
        NewCLIDCommand.Parameters.Add("@MatDocNo", SqlDbType.VarChar, 10, "MatDocNo")
        NewCLIDCommand.Parameters.Add("@MatDocItem", SqlDbType.VarChar, 10, "MatDocItem")
        NewCLIDCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 50, "Manufacturer")
        NewCLIDCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 50, "ManufacturerPN")
        NewCLIDCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
        NewCLIDCommand.Parameters.Add("@AddlData", SqlDbType.VarChar, 20, "AddlData")
        NewCLIDCommand.Parameters.Add("@Stemp", SqlDbType.VarChar, 50, "Stemp")
        NewCLIDCommand.Parameters.Add("@MSL", SqlDbType.VarChar, 50, "MSL")
        NewCLIDCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 50, "LastTransaction")

        NewCLIDCommand.Parameters("@OrgCode").Value = LoginData.OrgCode
        NewCLIDCommand.Parameters("@CreatedBy").Value = LoginData.User.ToUpper
        NewCLIDCommand.Parameters("@Operator").Value = LoginData.User.ToUpper
        NewCLIDCommand.Parameters("@Printed").Value = "True"
        NewCLIDCommand.Parameters("@IsTraceable").Value = "PT"
        NewCLIDCommand.Parameters("@AddlData").Value = SafeCode
        NewCLIDCommand.Parameters("@LastTransaction").Value = "MiscReceipt"

        myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.Parameters.AddWithValue("@NextNo", "")
        myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
        myCommand.Parameters(0).Size = 20
        myCommand.Parameters(0).Direction = ParameterDirection.Output
        myCommand.Parameters.AddWithValue("@TypeID", "CLID")
        'myCommand.Parameters.AddWithValue("@TypeID", "PCBID")
        myCommand.CommandTimeout = TimeOut_M5

        Dim lblPrint As String
        Dim RecDate As String
        Dim ProdDate As String

        ' Get valid GR Date                          Format: MM/dd/yyyy
        RecDate = Date.Now

        '' Get valid Receiving Date                  Format: MM/dd/yyyy
        'If Not LabelInfo.RecDate Is Nothing Then
        '    RecDate = Replace(LabelInfo.RecDate, "-", "/")
        'Else
        '    RecDate = ""
        'End If

        ' Get valid Expired Date                  Format: MM/dd/yyyy
        If Not LabelInfo.ExpDate Is Nothing Then
            ExpDate = Replace(LabelInfo.ExpDate, "-", "/")
        Else
            ExpDate = ""
        End If
        If LabelInfo.Units Is Nothing Then
            LabelInfo.Units = "1"
        End If

        Dim i, k As Integer
        'ErrorLogging("LabelForMiscReceipt", LoginData.User.ToUpper, "LabelInfo", "I")
        Try
            myConn.Open()

            For i = 0 To CInt(LabelInfo.Units) - 1

                'Try up to 5 times when failed getting next id
                NextCLID = ""
                k = 0
                While (k < 5 And NextCLID = "")
                    Try
                        ra = myCommand.ExecuteNonQuery()
                        NextCLID = myCommand.Parameters(0).Value
                    Catch ex As Exception
                        k = k + 1
                        ErrorLogging("LabelBasicFunction-LabelConversion", "Deadlocked? " & Str(k), "Failed getting next ID; RTNo: " & LabelInfo.RTNo & ", " & ex.Message & ex.Source, "E")
                    End Try
                End While

                If NextCLID <> "" Then

                    If PrintLabels = True Then     ' To ensure timely feedback, Do label printing later
                        lblPrint = "True"
                    Else
                        lblPrint = "Disabled"
                    End If

                    NewCLIDCommand.Parameters("@NewCLID").Value = NextCLID
                    NewCLIDCommand.Parameters("@MaterialNo").Value = LabelInfo.Material
                    NewCLIDCommand.Parameters("@MaterialRevision").Value = LabelInfo.MatlRev
                    NewCLIDCommand.Parameters("@Qty").Value = LabelInfo.Qty
                    NewCLIDCommand.Parameters("@UOM").Value = LabelInfo.ItemData.Tables(0).Rows(0)("UOM_CODE")
                    NewCLIDCommand.Parameters("@QtyBaseUOM").Value = LabelInfo.Qty
                    NewCLIDCommand.Parameters("@BaseUOM").Value = LabelInfo.ItemData.Tables(0).Rows(0)("UOM_CODE")
                    NewCLIDCommand.Parameters("@DateCode").Value = LabelInfo.DCode
                    NewCLIDCommand.Parameters("@LotNo").Value = LabelInfo.LotNo
                    NewCLIDCommand.Parameters("@COO").Value = DBNull.Value
                    NewCLIDCommand.Parameters("@RecDocNo").Value = LabelInfo.RTNo
                    NewCLIDCommand.Parameters("@RecDocItem").Value = DBNull.Value
                    NewCLIDCommand.Parameters("@RTLot").Value = LabelInfo.RTLot
                    NewCLIDCommand.Parameters("@VendorID").Value = ""
                    If Not RecDate = Nothing Then
                        NewCLIDCommand.Parameters("@RecDate").Value = CDate(RecDate)
                    Else
                        NewCLIDCommand.Parameters("@RecDate").Value = DBNull.Value
                    End If
                    If Not ExpDate = "" Then
                        NewCLIDCommand.Parameters("@ExpDate").Value = CDate(ExpDate)
                    Else
                        NewCLIDCommand.Parameters("@ExpDate").Value = DBNull.Value   'tmpDate
                    End If

                    If Not ProdDate = Nothing Then
                        NewCLIDCommand.Parameters("@ProdDate").Value = CDate(ProdDate)
                    Else
                        NewCLIDCommand.Parameters("@ProdDate").Value = DBNull.Value
                    End If
                    NewCLIDCommand.Parameters("@RoHS").Value = LabelInfo.RoHS               '"COMPLIANT"         
                    NewCLIDCommand.Parameters("@PurOrdNo").Value = ""
                    NewCLIDCommand.Parameters("@PurOrdItem").Value = ""
                    NewCLIDCommand.Parameters("@InvoiceNo").Value = ""
                    NewCLIDCommand.Parameters("@BillofLading").Value = ""
                    NewCLIDCommand.Parameters("@DN").Value = ""
                    NewCLIDCommand.Parameters("@HeaderText").Value = ""
                    NewCLIDCommand.Parameters("@ReasonCode").Value = ""

                    If LabelInfo.ItemData.Tables(0).Rows(0)("ROUTING_ID").ToString = "2" Then
                        NewCLIDCommand.Parameters("@StockType").Value = "QP"
                    Else
                        NewCLIDCommand.Parameters("@StockType").Value = "FTS"
                    End If

                    NewCLIDCommand.Parameters("@MaterialDesc").Value = SQLString(LabelInfo.ItemData.Tables(0).Rows(0)("ITEM_DESC"))
                    NewCLIDCommand.Parameters("@VendorName").Value = ""
                    NewCLIDCommand.Parameters("@VendorPN").Value = ""
                    NewCLIDCommand.Parameters("@SLOC").Value = LabelInfo.SubInv
                    NewCLIDCommand.Parameters("@StorageBin").Value = LabelInfo.Locator
                    NewCLIDCommand.Parameters("@MatDocNo").Value = LabelInfo.RTNo              'Add by Yudy 09/08/2008		
                    NewCLIDCommand.Parameters("@MatDocItem").Value = DBNull.Value
                    NewCLIDCommand.Parameters("@Manufacturer").Value = LabelInfo.MFR
                    NewCLIDCommand.Parameters("@ManufacturerPN").Value = LabelInfo.MPN
                    NewCLIDCommand.Parameters("@QMLStatus").Value = DBNull.Value
                    NewCLIDCommand.Parameters("@Stemp").Value = LabelInfo.Stemp
                    NewCLIDCommand.Parameters("@MSL").Value = LabelInfo.MSL

                    NewCLIDCommand.CommandTimeout = TimeOut_M5
                    NewCLIDCommand.CommandType = CommandType.Text
                    ra = NewCLIDCommand.ExecuteNonQuery()

                    myDataRow = CLIDsTable.NewRow()
                    myDataRow("CLID") = NextCLID
                    CLIDsTable.Rows.Add(myDataRow)
                End If
            Next
            LabelForMiscReceipt.CLIDs = New DataSet
            LabelForMiscReceipt.CLIDs = CLIDs

            myConn.Close()
        Catch ex As Exception
            ErrorLogging("LabelForMiscReceipt", LoginData.User.ToUpper, "RTNo: " & LabelInfo.RTNo & ", " & ex.Message & ex.Source, "E")
            LabelForMiscReceipt.ErrorMsg = "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ExportHWDataInfo(ByVal User As String, ByVal exportdata As String) As HW_ExportDataInfo
        Try
            Dim xmlReader As New XmlDocument
            xmlReader.LoadXml(exportdata)
            Dim iCycle, iNode, iResult, iError As Integer
            Dim result As HW_ExportDataInfo
            For iCycle = 0 To xmlReader.DocumentElement.ChildNodes.Count - 1
                If xmlReader.DocumentElement.ChildNodes(iCycle).Name.ToUpper.Trim = "Export".ToUpper.Trim Then
                    For iNode = 0 To xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes.Count - 1
                        If xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).Name.ToUpper.Trim = "ResultData".ToUpper.Trim Then
                            For iResult = 0 To xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes.Count - 1
                                If xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "TransSourceID".ToUpper.Trim Then
                                    result.TransSourceID = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "BarCode".ToUpper.Trim Then
                                    result.BarCode = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "SnNO".ToUpper.Trim Then
                                    result.SnNO = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Comments".ToUpper.Trim Then
                                    result.Comments = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "CreatedDate".ToUpper.Trim Then
                                    result.CreatedDate = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "CreatedBy".ToUpper.Trim Then
                                    result.CreatedBy = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Updated_by".ToUpper.Trim Then
                                    result.Updated_by = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Updated_date".ToUpper.Trim Then
                                    result.Updated_date = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment1".ToUpper.Trim Then
                                    result.Segment1 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment2".ToUpper.Trim Then
                                    result.Segment2 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment3".ToUpper.Trim Then
                                    result.Segment3 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment4".ToUpper.Trim Then
                                    result.Segment4 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment5".ToUpper.Trim Then
                                    result.Segment5 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment6".ToUpper.Trim Then
                                    result.Segment6 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment7".ToUpper.Trim Then
                                    result.Segment7 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment8".ToUpper.Trim Then
                                    result.Segment8 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).Name.ToUpper.Trim = "Segment9".ToUpper.Trim Then
                                    result.Segment9 = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iResult).InnerText.Trim
                                End If
                            Next
                        End If
                        If xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).Name.ToUpper.Trim = "Message".ToUpper.Trim Then
                            For iError = 0 To xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes.Count - 1
                                If xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iError).Name.ToUpper.Trim = "ErrorCode".ToUpper.Trim Then
                                    result.ErrorCode = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iError).InnerText.Trim
                                ElseIf xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iError).Name.ToUpper.Trim = "ErrorMsg".ToUpper.Trim Then
                                    result.ErrorMsg = xmlReader.DocumentElement.ChildNodes(iCycle).ChildNodes(iNode).ChildNodes(iError).InnerText.Trim
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            Return result
        Catch ex As Exception
            ErrorLogging("ExportHWDataInfo", User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Exit Function
        End Try
    End Function


    Public Function Save_SubInvTransfer(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv    '53485     '  CAROLD3  53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Dim flag As Boolean
        Dim Check_VMIList, DstInv, SrcInv, trans_type As String
        Dim Check_DstInv, Check_SrcInv As Integer

        Try
            Check_VMIList = Convert.ToString(da.ExecuteScalar(String.Format("select Value from T_Config with (nolock) where ConfigID = 'CLID013'")))
            If Check_VMIList Is Nothing OrElse Check_VMIList Is DBNull.Value OrElse Check_VMIList.ToString = "" Then

            Else
                DstInv = p_ds.Tables("transfer_table").Rows(0)("p_subinventory_destination")
                Check_DstInv = InStr(Check_VMIList, DstInv)
                If Check_DstInv > 0 Then
                    SrcInv = p_ds.Tables("transfer_table").Rows(0)("p_subinventory_source")
                    Check_SrcInv = InStr(Check_VMIList, SrcInv)
                    If Check_SrcInv > 0 Then

                    Else
                        p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                        p_ds.Tables("transfer_table").Rows(0)("o_return_message") = "Can't direct transfer back to VMI subinventory from inhouse"
                        ErrorLogging(MoveType & "CheckVMIFlag", OracleLoginData.User, "Can't direct transfer back to VMI subinventory " & DstInv & " from " & SrcInv, "I")
                        Return p_ds
                    End If
                Else

                End If
            End If
        Catch ex As Exception
            p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "CheckVMIFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return p_ds
        End Try

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            'If MoveType = "MMt-SubTran" Then
            '    trans_type = "Subinventory Transfer"
            'ElseIf MoveType = "MMt-TrmSub" Then
            '    trans_type = "Trimming Replenishment"
            'ElseIf MoveType = "MMt-SubScr" Then
            '    trans_type = "Scrap Materials"
            'ElseIf MoveType = "MMt-WrtOff" Then
            '    trans_type = "Transfer Prewritoff(E&O)"
            'ElseIf MoveType = "MMt-BckFlsh" Then
            '    trans_type = "Backflush Transfer"
            'ElseIf MoveType = "MMt-SubOSP" Then
            '    trans_type = "Move OSP Supplier Buffer Stock"
            'ElseIf MoveType = "MMt-PrdRtn" Then
            '    trans_type = "Prod Return for Trimming"
            'End If

            'ErrorLogging(MoveType & "-Post", OracleLoginData.User, "Mark", "I")

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.subinventory_transfer_batch2"
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
            'comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240).Value = trans_type
            comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_chargeable", OracleType.VarChar, 240)  ' --old parameter -->Scrap Notice
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
            comm.Parameters("p_chargeable").SourceColumn = "p_chargeable"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("transfer_table"))

            Dim i As Integer
            Dim result_flag As String
            Dim DR() As DataRow = Nothing
            DR = p_ds.Tables("transfer_table").Select("o_return_status <> 'Y'")

            If DR.Length = 0 Then
                result_flag = submit_inv_cvs("subinventory_transfer_batch2", TransactionID, OracleLoginData)

                If result_flag = "Y" Then
                    Try
                        flag = eTrace_Update(UpdateTable, OracleLoginData, MoveType)
                    Catch ex As Exception
                        p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                        p_ds.Tables("transfer_table").Rows(0)("o_return_message") = "eTrace update error"
                        ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, ex.Message.ToString, "E")
                        Return p_ds
                        'Throw ex
                    End Try
                    If flag = True Then
                        Return p_ds
                    Else
                        p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                        p_ds.Tables("transfer_table").Rows(0)("o_return_message") = "eTrace update error"
                        ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, "eTrace update error", "I")
                        Return p_ds
                    End If
                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag, "I")
                    p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("transfer_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_inv_cvs("subinventory_transfer_batch2", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next
                result_flag = del_inv_cvs("subinventory_transfer_batch2", TransactionID, OracleLoginData)
                Return p_ds
            End If

            'If DR.Length = 0 Then
            '    'ErrorLogging(MoveType, OracleLoginData.User, p_ds.Tables("transfer_table").Rows(0)("p_subinventory_destination"))
            '    comm.Transaction.Commit()
            '    comm.Connection.Close()
            '    'comm.Connection.Dispose()
            '    'comm.Dispose()
            '    Try
            '        flag = eTrace_Update(UpdateTable, OracleLoginData, MoveType)
            '    Catch ex As Exception
            '        p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            '        p_ds.Tables("transfer_table").Rows(0)("o_return_message") = "eTrace update error"
            '        ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, ex.Message.ToString)
            '        Return p_ds
            '        'Throw ex
            '    End Try
            '    If flag = True Then
            '        Return p_ds
            '    Else
            '        p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            '        p_ds.Tables("transfer_table").Rows(0)("o_return_message") = "eTrace update error"
            '        ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, "eTrace update error")
            '        Return p_ds
            '    End If
            'Else
            '    For i = 0 To (DR.Length - 1)
            '        ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString)
            '    Next
            '    'For Each dtRow In DR
            '    '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString)
            '    'Next

            '    'comm.Transaction.Rollback()
            '    comm.Transaction.Commit()
            '    comm.Connection.Close()
            '    'comm.Connection.Dispose()
            '    'comm.Dispose()
            '    Return p_ds
            'End If
        Catch ex As Exception
            p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            del_inv_cvs("subinventory_transfer_batch2", TransactionID, OracleLoginData)
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function save_account_receipt(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv  '53485   'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv   '401
        Dim flag As Boolean

        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_receipt"      '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
            comm.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountreceipt_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountreceipt_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Try
                    flag = eTrace_Update(UpdateTable, OracleLoginData, MoveType)
                Catch ex As Exception
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = "eTrace update error"
                    ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, ex.Message & ex.Source, "E")
                    Return p_ds
                    'Throw ex
                End Try
                If flag = True Then
                    Return p_ds
                Else
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = "eTrace update error"
                    ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, "eTrace update error", "I")
                    Return p_ds
                End If
                'Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            'comm.Transaction.Rollback()
            comm.Transaction.Commit()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function save_account_issue(ByVal p_ds As DataSet, ByVal UpdateTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
        Dim appl As Integer = OracleLoginData.AppID_Inv    '401
        Dim flag As Boolean
        Try
            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.account_alias_issue"   '53330   401
            comm.Parameters.Add("p_timeout", OracleType.Double)
            comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
            comm.Parameters.Add("p_transaction_header_id", OracleType.Double)
            comm.Parameters.Add("p_transaction_source_name", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
            comm.Parameters.Add("p_source_line_id", OracleType.Double)
            comm.Parameters.Add("p_source_header_id", OracleType.Double)
            comm.Parameters.Add("p_user_id", OracleType.Double)
            comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240)
            comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240)
            comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240)
            comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
            comm.Parameters.Add("p_transaction_quantity", OracleType.Double)
            comm.Parameters.Add("p_primary_quantity", OracleType.Double)
            comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
            comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
            comm.Parameters.Add("p_scrap_notice", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
            comm.Parameters.Add("o_return_message", OracleType.VarChar, 240)

            comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

            comm.Parameters("p_timeout").SourceColumn = "p_timeout"
            'comm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
            comm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
            comm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
            comm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
            comm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
            comm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
            comm.Parameters("p_user_id").SourceColumn = "p_user_id"
            comm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
            comm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
            comm.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
            comm.Parameters("p_locator_source").SourceColumn = "p_locator_source"
            comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
            comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("p_scrap_notice").SourceColumn = "p_scrap_notice"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm
            oda_h.Update(p_ds.Tables("accountissue_table"))

            Dim DR() As DataRow = Nothing
            'Dim dtRow As DataRow
            Dim i As Integer
            DR = p_ds.Tables("accountissue_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                comm.Transaction.Commit()
                comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                Try
                    flag = eTrace_Update(UpdateTable, OracleLoginData, MoveType)
                Catch ex As Exception
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = "eTrace update error"
                    ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, ex.Message & ex.Source, "E")
                    Return p_ds
                    'Throw ex
                End Try
                If flag = True Then
                    Return p_ds
                Else
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = "eTrace update error"
                    ErrorLogging(MoveType & "eTrace_Update", OracleLoginData.User, "eTrace update error", "I")
                    Return p_ds
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString, "I")
                'Next

                'comm.Transaction.Rollback()
                comm.Transaction.Commit()
                comm.Connection.Close()
                Return p_ds
            End If

        Catch ex As Exception
            p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging(MoveType & "-Post", OracleLoginData.User, ex.Message & ex.Source, "E")
            'comm.Transaction.Rollback()
            comm.Transaction.Commit()
            Return p_ds
            Throw ex
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function eTrace_Update(ByVal UpdtTable As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String) As Boolean

        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim flag As Integer
        Dim ra As Integer
        Dim i As Integer
        Dim FromSubinv, ToSubinv, sqlCMD, StockType As String

        Try
            If MoveType = "MMt-SubTran" Then
                Using da As DataAccess = GetDataAccess()
                    sqlCMD = String.Format("select SLOC from T_CLMaster with (nolock) where CLID = '{0}'", UpdtTable.Tables(0).Rows(0)("CLID"))
                    FromSubinv = FixNull(da.ExecuteScalar(sqlCMD))

                    ToSubinv = UpdtTable.Tables(0).Rows(0)("p_transfer_subinventory")

                    StockType = FixNull(da.ExecuteScalar(String.Format("select StockType from T_LabelImprove with (nolock) where OrgCode = '{0}' and FromSubInv = '{1}' and ToSubInv = '{2}'", OracleLoginData.OrgCode, FromSubinv, ToSubinv)))
                End Using
            End If

            For i = 0 To UpdtTable.Tables(0).Rows.Count - 1
                If MoveType = "MMt-SubTran" AndAlso StockType <> "" Then
                    CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM = @QtyBaseUOM,ChangedOn = getdate(),ChangedBy = @ChangedBy,StatusCode = @StatusCode,ReasonCode = @ReasonCode,SLOC = @SLOC,StorageType = @StorageType, StorageBin = @StorageBin,StockType = @StockType, LastTransaction = @LastTransaction where CLID = @CLID", myConn)
                    CLMasterSQLCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 20).Value = StockType
                Else
                    CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM = @QtyBaseUOM,ChangedOn = getdate(),ChangedBy = @ChangedBy,StatusCode = @StatusCode,ReasonCode = @ReasonCode,SLOC = @SLOC,StorageType = @StorageType, StorageBin = @StorageBin,LastTransaction = @LastTransaction where CLID = @CLID", myConn)
                End If

                CLMasterSQLCommand.Parameters.Add("@QtyBaseUOM", SqlDbType.VarChar, 24, "QtyBaseUOM")
                CLMasterSQLCommand.Parameters.Add("@ChangedBy", SqlDbType.VarChar, 50, "ChangedBy")
                CLMasterSQLCommand.Parameters.Add("@StatusCode", SqlDbType.VarChar, 50, "StatusCode")
                CLMasterSQLCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
                CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
                CLMasterSQLCommand.Parameters.Add("@StorageType", SqlDbType.VarChar, 20, "StorageType")
                CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
                CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
                CLMasterSQLCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 50, "LastTransaction")

                CLMasterSQLCommand.Parameters("@QtyBaseUOM").Value = UpdtTable.Tables(0).Rows(i)("p_qty")
                CLMasterSQLCommand.Parameters("@ChangedBy").Value = OracleLoginData.User.ToUpper
                CLMasterSQLCommand.Parameters("@StatusCode").Value = UpdtTable.Tables(0).Rows(i)("p_status")
                CLMasterSQLCommand.Parameters("@ReasonCode").Value = UpdtTable.Tables(0).Rows(i)("p_reasoncode")
                CLMasterSQLCommand.Parameters("@SLOC").Value = UpdtTable.Tables(0).Rows(i)("p_transfer_subinventory")
                CLMasterSQLCommand.Parameters("@StorageType").Value = UpdtTable.Tables(0).Rows(i)("p_storagetype")
                CLMasterSQLCommand.Parameters("@StorageBin").Value = UpdtTable.Tables(0).Rows(i)("p_destlocator")
                CLMasterSQLCommand.Parameters("@CLID").Value = UpdtTable.Tables(0).Rows(i)("CLID")
                CLMasterSQLCommand.Parameters("@LastTransaction").Value = MoveType
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5

                Try
                    myConn.Open()
                    CLMasterSQLCommand.CommandType = CommandType.Text
                    ra = CLMasterSQLCommand.ExecuteNonQuery()
                    myConn.Close()
                    flag = 0
                Catch ex As Exception
                    ErrorLogging(MoveType & "-eTrace_Update", OracleLoginData.User, "CLID:" & UpdtTable.Tables(0).Rows(i)("CLID") & " update to Qty:" & UpdtTable.Tables(0).Rows(i)("p_qty") & " SubInv/Loc:" & UpdtTable.Tables(0).Rows(i)("p_transfer_subinventory") & "/" & UpdtTable.Tables(0).Rows(i)("p_destlocator") & ". With error: " & ex.Message & ex.Source, "E")
                    flag = flag + 1
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try
            Next

            If MoveType = "MMt-CostSub" AndAlso Mid(UpdtTable.Tables(0).Rows(0)("CLID"), 1, 1) = "B" Then
                Dim Update_Rst As Boolean
                Using da As DataAccess = GetDataAccess()
                    Update_Rst = da.ExecuteNonQuery(String.Format("exec sp_Update_Shippment_ForMisRcpt N'{0}',N'{1}'", DStoXML(UpdtTable), OracleLoginData.User))
                End Using
            End If

            If flag > 0 Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            ErrorLogging("eTrace_Update", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ChangeCLID(ByVal CLID As String, ByVal Qty As Double, ByVal DestSub As String, ByVal DestLoc As String, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal SlotNo As String) As Boolean
        Dim flag As Integer
        Dim clidQty, ReqQty, PickedQty, CurQty, DifQty, OldPickQty As Decimal
        Dim OrgCode, Item, PDTO, Status, SLOC, LastDJ As String
        Dim dsPick, dsCLID, dsPO_CLID As DataSet
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strCMD As String

        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                dsCLID = New DataSet
                sqlstr = String.Format("SELECT OrgCode, MaterialNo, QtyBaseUOM, SLOC, LastDJ FROM T_CLMaster with (nolock) WHERE CLID='{0}'", CLID)
                dsCLID = da.ExecuteDataSet(sqlstr, "CLID")

                OrgCode = FixNull(dsCLID.Tables(0).Rows(0)("OrgCode"))
                clidQty = dsCLID.Tables(0).Rows(0)("QtyBaseUOM")
                Item = dsCLID.Tables(0).Rows(0)("MaterialNo")
                SLOC = FixNull(dsCLID.Tables(0).Rows(0)("SLOC"))
                LastDJ = FixNull(dsCLID.Tables(0).Rows(0)("LastDJ"))
            Catch ex As Exception
                flag = flag + 1
                ErrorLogging("MMC-ChangeCLID", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using

        'If lastdj <> "" Then
        '    strCMD = String.Format("SELECT PDTO FROM T_PO_CLID WHERE PDTO IS NOT NULL and PO = '{1}' and CLID='{0}' and ReturnDate is NULL", CLID, LastDJ)
        'Else
        '    strCMD = String.Format("SELECT PDTO FROM T_PO_CLID WHERE PDTO IS NOT NULL and CLID='{0}' and ReturnDate is NULL", CLID)
        'End If
        'Try
        '    myConn.Open()
        '    cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
        '    cmdSQL.CommandTimeout = TimeOut_M5
        '    PDTO = Convert.ToString(cmdSQL.ExecuteScalar).Trim
        '    myConn.Close()
        '    flag = 0
        'Catch ex As Exception
        '    If myConn.State <> ConnectionState.Closed Then
        '        myConn.Close()
        '    End If
        '    ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Select PDTO from CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
        '    flag = flag + 1
        'End Try

        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                dsPO_CLID = New DataSet
                sqlstr = String.Format("SELECT CLIDQty FROM T_PO_CLID with (nolock) WHERE IssueDate=(select max(IssueDate) from T_PO_CLID with (nolock) where CLID='{0}') and CLID='{0}'", CLID)
                dsPO_CLID = da.ExecuteDataSet(sqlstr, "PO_CLID")

                If Not dsPO_CLID Is Nothing AndAlso dsPO_CLID.Tables.Count > 0 AndAlso dsPO_CLID.Tables(0).Rows.Count > 0 Then
                    OldPickQty = dsPO_CLID.Tables(0).Rows(0)("CLIDQty")
                Else
                    OldPickQty = 0
                End If
            Catch ex As Exception
                flag = flag + 1
                ErrorLogging("MMC-ChangeCLID_GetQty", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using

        'If lastdj <> "" Then
        '    strCMD = String.Format("UPDATE T_PO_CLID set CLIDQty={0}-{1}, ReturnDate=getDate(), ChangedBy='{2}' from T_PO_CLID where PO = '{3}' and CLID='{4}' and ReturnDate is NULL", clidQty, Qty, OracleLoginData.User, lastdj, CLID)
        'Else
        strCMD = String.Format("UPDATE T_PO_CLID set CLIDQty={0}-{1}, ReturnDate=getDate(), ChangedBy='{2}' where IssueDate=(select max(IssueDate) from T_PO_CLID where CLID='{3}' and ReturnDate is NULL)  and  CLID='{3}' and ReturnDate is NULL ", OldPickQty, Qty, OracleLoginData.User, CLID)
        'End If
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
            cmdSQL.CommandTimeout = TimeOut_M5
            cmdSQL.ExecuteNonQuery()
            myConn.Close()
            flag = 0
        Catch ex As Exception
            ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Update T_PO_CLID with CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
            flag = flag + 1
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        'If Not PDTO Is Nothing AndAlso FixNull(PDTO) <> "" Then
        '    Using da As DataAccess = GetDataAccess()
        '        Try
        '            Dim sqlstr, OldRemark As String
        '            dsPick = New DataSet
        '            sqlstr = String.Format("SELECT ReqQty, PickedQty, Status,Remark FROM T_PDTOItem WHERE PDTO='{0}' and Material = '{1}'", PDTO, Item)
        '            dsPick = da.ExecuteDataSet(sqlstr, "PickData")

        '            ReqQty = dsPick.Tables(0).Rows(0)("ReqQty")
        '            PickedQty = dsPick.Tables(0).Rows(0)("PickedQty")
        '            Status = dsPick.Tables(0).Rows(0)("Status")
        '            CurQty = PickedQty - Qty
        '            DifQty = ReqQty - CurQty
        '            OldRemark = FixNull(dsPick.Tables(0).Rows(0)("Remark"))

        '            ErrorLogging("MMC-ChangeCLID", OracleLoginData.User.ToUpper, "PickedQty: " & PickedQty & " , ReturnQty: " & Qty & " , Old Remark: " & OldRemark, "I")
        '        Catch ex As Exception
        '            flag = flag + 1
        '            ErrorLogging("MMC-ChangeCLID", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        '        End Try
        '    End Using

        '    Dim Remark As String = DateTime.Now.ToString & " by User: " & OracleLoginData.User.ToUpper & " with LastTransaction: ChangeCLID"
        '    If Status <> 0 Then
        '        strCMD = String.Format("UPDATE T_PDTOItem set PickedQty={0}, Remark = '{3}' where PDTO = '{1}' and Material = '{2}'", CurQty, PDTO, Item, Remark)
        '    ElseIf Status = 0 Then
        '        If DifQty > 0 Then
        '            strCMD = String.Format("UPDATE T_PDTOItem set PickedQty={0}, Status = '1', Remark = '{3}' where PDTO = '{1}' and Material = '{2}'", CurQty, PDTO, Item, Remark)
        '        Else
        '            strCMD = String.Format("UPDATE T_PDTOItem set PickedQty={0}, Remark = '{3}' where PDTO = '{1}' and Material = '{2}'", CurQty, PDTO, Item, Remark)
        '        End If
        '    End If
        '    Try
        '        myConn.Open()
        '        cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
        '        cmdSQL.CommandTimeout = TimeOut_M5
        '        cmdSQL.ExecuteNonQuery()
        '        myConn.Close()
        '        flag = 0
        '    Catch ex As Exception
        '        ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Update T_PDTOItem for PDTO: " & PDTO & " and item" & Item & ", " & ex.Message & ex.Source, "E")
        '        flag = flag + 1
        '    Finally
        '        If myConn.State <> ConnectionState.Closed Then myConn.Close()
        '    End Try
        'End If

        If DestSub <> "" Then
            Using da As DataAccess = GetDataAccess()
                Try
                    Dim Sqlstr As String
                    Dim Cfg, SubInv, rs As String

                    'ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Start inserting record to T_PO_CLID_RTN with CLID: " & CLID, "I")

                    Sqlstr = String.Format("select Value from T_Config with(nolock) where ConfigID = 'MMT011'")
                    Cfg = da.ExecuteScalar(Sqlstr).ToString

                    If Cfg = "YES" Then
                        Sqlstr = String.Format("Select Description from T_SysLov where Name = 'BF Count'  and Description = '{0}'", SLOC)
                        SubInv = FixNull(da.ExecuteScalar(Sqlstr))

                        If SubInv <> "" Then
                            Sqlstr = String.Format("Insert into T_PO_CLID_RTN Values('{0}', getdate(), '{1}', '{2}', '', '{3}')", CLID, Qty, OrgCode, OracleLoginData.User)
                            rs = da.ExecuteNonQuery(Sqlstr)

                            'ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Finish inserting record to T_PO_CLID_RTN with CLID: " & CLID, "I")
                        End If
                    Else

                    End If

                Catch ex As Exception
                    ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Insert record to T_PO_CLID_RTN with CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                    flag = flag + 1
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try
            End Using
        End If

        If DestSub = "" Then
            'strCMD = String.Format("UPDATE T_CLMaster SET StatusCode=(case when StatusCode='0' then '1' else '2' end),QtyBaseUOM={0},ChangedOn=getDate(),ChangedBy='{1}',SupplyType='',LastDJ='' WHERE CLID='{2}'", Qty, OracleLoginData.User, CLID)
            strCMD = String.Format("UPDATE T_CLMaster SET StatusCode= '1',QtyBaseUOM={0},ChangedOn=getDate(),ChangedBy='{1}',StorageType = '{4}', SupplyType='',LastDJ='',LastTransaction ='{3}' WHERE CLID='{2}'", Qty, OracleLoginData.User, CLID, MoveType, SlotNo)
        Else
            'strCMD = String.Format("UPDATE T_CLMaster SET StatusCode=(case when StatusCode='0' then '1' else '2' end),QtyBaseUOM={0},ChangedOn=getDate(),ChangedBy='{1}',SLOC='{3}',StorageBin='{4}',SupplyType='',LastDJ='' WHERE CLID='{2}'", Qty, OracleLoginData.User, CLID, DestSub, DestLoc)
            strCMD = String.Format("UPDATE T_CLMaster SET StatusCode= '1',QtyBaseUOM={0},ChangedOn=getDate(),ChangedBy='{1}',SLOC='{3}',StorageType = '{6}', StorageBin='{4}',SupplyType='',LastDJ='',LastTransaction ='{5}' WHERE CLID='{2}'", Qty, OracleLoginData.User, CLID, DestSub, DestLoc, MoveType, SlotNo)
        End If
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
            cmdSQL.CommandTimeout = TimeOut_M5
            cmdSQL.ExecuteNonQuery()
            myConn.Close()
            flag = 0
        Catch ex As Exception
            ErrorLogging(MoveType, OracleLoginData.User.ToUpper, "Update T_CLMaster with CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
            flag = flag + 1
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        If flag > 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function UpdateAML(ByVal OracleLoginData As ERPLogin, ByVal AMLData As DataSet) As DataSet
        Dim I As Integer
        Dim cmdSQL As SqlClient.SqlCommand
        Dim strCMD As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            myConn.Open()
            For I = 0 To AMLData.Tables("AMLTable").Rows.Count - 1
                strCMD = String.Format("UPDATE T_CLMaster set Manufacturer='{0}', ManufacturerPN='{1}' where CLID='{2}' and OrgCode='{3}'", AMLData.Tables("AMLTable").Rows(I)("MFR"), AMLData.Tables("AMLTable").Rows(I)("MPN"), AMLData.Tables("AMLTable").Rows(I)("CLID"), OracleLoginData.OrgCode)
                cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
                cmdSQL.ExecuteNonQuery()
            Next
            myConn.Close()
        Catch ex As Exception
            If myConn.State <> ConnectionState.Closed Then
                myConn.Close()
            End If
            ErrorLogging("AML_Update", OracleLoginData.User.ToUpper, ex.Message)
        Finally
            If myConn.State <> ConnectionState.Closed Then
                myConn.Close()
            End If
        End Try

    End Function

    Public Function GetRTNo(ByVal OracleLoginData As ERPLogin) As String
        Dim RTNo As String
        RTNo = CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))
        Return RTNo
    End Function

    Public Function GetRTNo_MiscRcpt(ByVal OracleLoginData As ERPLogin) As String
        Dim RTNo As String
        RTNo = CLng(Format(DateTime.Now, "yyyyMMdd"))
        Return RTNo
    End Function

    Public Function GetDate(ByVal OracleLoginData As ERPLogin) As Date
        Return DateTime.Now
    End Function

    Public Function GetNextInvID(ByVal OracleLoginData As ERPLogin) As String
        Dim NextInvID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 7
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "INVID")
            'myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            'myCommand.CommandType = CommandType.StoredProcedure
            'myCommand.Parameters.AddWithValue("@NextNo", "")
            'myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            'myCommand.Parameters(0).Size = 20
            'myCommand.Parameters(0).Direction = ParameterDirection.Output
            'myCommand.Parameters.AddWithValue("@TypeID", "CLID")

            'Try up to 5 times when failed getting next id
            NextInvID = ""
            k = 0
            While (k < 5 And NextInvID = "")
                Try
                    myCommand.CommandTimeout = TimeOut_M5
                    ra = myCommand.ExecuteNonQuery()
                    NextInvID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextInvID", "Deadlocked? " & Str(k), "Failed to get next InvID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextInvID
        Catch ex As Exception
            ErrorLogging("GetNextInvID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextInvID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetNextHeaderID(ByVal OracleLoginData As ERPLogin) As String
        Dim NextInvID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 9
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "HDID")
            'myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            'myCommand.CommandType = CommandType.StoredProcedure
            'myCommand.Parameters.AddWithValue("@NextNo", "")
            'myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            'myCommand.Parameters(0).Size = 20
            'myCommand.Parameters(0).Direction = ParameterDirection.Output
            'myCommand.Parameters.AddWithValue("@TypeID", "CLID")

            'Try up to 5 times when failed getting next id
            NextInvID = ""
            k = 0
            While (k < 5 And NextInvID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextInvID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextHeaderID", "Deadlocked? " & Str(k), "Failed to get next HeaderID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextInvID
        Catch ex As Exception
            ErrorLogging("GetNextHeaderID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextInvID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetNextCLID(ByVal OracleLoginData As ERPLogin) As String
        Dim NextCLID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "CLID")
            myCommand.CommandTimeout = TimeOut_M5
            'Try up to 5 times when failed getting next id
            k = 0
            While (k < 5 And NextCLID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextCLID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextCLID", "Deadlocked? " & Str(k), "Failed to get next InvID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextCLID
        Catch ex As Exception
            ErrorLogging("GetNextCLID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextCLID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetNextProdID(ByVal OracleLoginData As ERPLogin) As String
        Dim NextProdID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "PCBID")
            myCommand.CommandTimeout = TimeOut_M5
            'Try up to 5 times when failed getting next id
            k = 0
            While (k < 5 And NextProdID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextProdID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextProdID", "Deadlocked? " & Str(k), "Failed to get next InvID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextProdID
        Catch ex As Exception
            ErrorLogging("GetNextProdID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextProdID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function CheckFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim i As Integer
        Dim DR_Blank() As DataRow = Nothing

        Try
            'deal with blank rows
            DR_Blank = ExcelData.Tables(0).Select("(item is null or item = '') and (unitqty is null or unitqty = '') and (noofpackage is null or noofpackage = '') and (subinventory is null or subinventory = '')")
            If DR_Blank.Length > 0 Then
                For Each drblank As DataRow In DR_Blank
                    ExcelData.Tables(0).Rows.Remove(drblank)
                Next
            End If

            For Each dr As DataRow In ExcelData.Tables(0).Rows
                'deal with blank rows
                'If FixNull(dr(0)) = "" Then
                '    ExcelData.Tables(0).Rows.Remove(dr)
                '    Continue For
                'End If

                RowIndex = RowIndex + 1
                myDataRow = MRListData.Tables("MRData").NewRow()
                'myDataRow("Select") = dr(0).ToString.Trim
                myDataRow("InvID") = InvID
                myDataRow("Line") = RowIndex
                myDataRow("Item") = UCase(IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim))
                myDataRow("Revision") = UCase(IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim))
                'myDataRow("UnitQty") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                'myDataRow("UnitQty") = IIf(dr(2) Is DBNull.Value, 0, FormatNumber(dr(2), 3))
                If FixNull(dr(2)) = "" OrElse FixNull(dr(2)) = 0 Then
                    myDataRow("UnitQty") = 0
                Else
                    myDataRow("UnitQty") = CDec(Format(dr(2), "#.######"))  'FormatNumber(dr(2), 6)
                End If
                myDataRow("NoofPackage") = IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim)
                myDataRow("Subinventory") = UCase(IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim))
                myDataRow("Locator") = UCase(IIf(dr(5) Is DBNull.Value, "", dr(5).ToString.Trim))
                If myDataRow("Locator") <> "" Then
                    myDataRow("Locator") = Replace(myDataRow("Locator"), "…", "...")
                End If
                myDataRow("Manufacturer") = UCase(IIf(dr(6) Is DBNull.Value, "", CStr(FixNull(dr(6))).Trim))
                myDataRow("ManufacturerPN") = UCase(IIf(dr(7) Is DBNull.Value, "", CStr(FixNull(dr(7))).Trim))
                myDataRow("ExpiredDate") = IIf(dr(8) Is DBNull.Value, "", dr(8).ToString.Trim)
                'myDataRow("ExpiredDate") = IIf(Fix(dr(8)) = "", "", Date.TryParse(dr(8), dt).ToString.Trim)
                myDataRow("DateCode") = IIf(dr(9) Is DBNull.Value, "", CStr(FixNull(dr(9))).Trim)
                myDataRow("LotNo") = IIf(dr(10) Is DBNull.Value, "", CStr(FixNull(dr(10))).Trim)

                myDataRow("BPCSPN") = UCase(IIf(dr(11) Is DBNull.Value, "", CStr(FixNull(dr(11))).Trim))
                'If Not dr(10) Is DBNull.Value Then myDataRow("ManufactureDate") = CDate(dr(10))
                If FixNull(dr(0)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Item must be entered"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(2)) = "" Then
                    myDataRow("ErrorMsg") = "Error: UnitQty must be entered"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(2)) = "" AndAlso dr(2) = 0 Then
                    myDataRow("ErrorMsg") = "Error: UnitQty can't be Zero"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(2)) = "" AndAlso dr(2) < 0.00001 Then
                    myDataRow("ErrorMsg") = "Error: UnitQty can't be less than 0.00001"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(3)) = "" Then
                    myDataRow("ErrorMsg") = "Error: NoofPackage must be entered"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(4)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Subinventory must be entered"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If

                If Double.TryParse(dr(2), d) = False OrElse dr(2) <= 0 Then
                    myDataRow("ErrorMsg") = "Error: Enter UnitQty only with Positive number"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If

                If dr(3) <= 0 Then
                    myDataRow("ErrorMsg") = "Error: Enter NoOfPackage only with Positive Integer"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf dr(3) > 0 And dr(3) <> Int(dr(3)) Then
                    myDataRow("ErrorMsg") = "Error: Enter NoOfPackage only with Integer"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If

                If FixNull(dr(8)) <> "" AndAlso Date.TryParse(dr(8), dt) = False Then
                    myDataRow("ErrorMsg") = "Error: Enter ExpiredDate with format: MM/DD/YYYY, do Exceed Today"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(8)) <> "" AndAlso Date.TryParse(dr(8), dt) = True Then
                    Dim datenow As Date
                    Dim DifOfDate As Integer
                    datenow = GetDate(OracleERPLogin)
                    DifOfDate = DateDiff("d", datenow, dr(8))
                    If DifOfDate < 1 Or Year(dr(8)) > Year(datenow) + 5 Then
                        myDataRow("ErrorMsg") = "Error: Exp Date " & dr(8) & " is not in correct interval"
                        MRListData.Tables("MRData").Rows.Add(myDataRow)
                        Continue For
                    End If
                End If
                If FixNull(dr(8)) <> "" Then
                    myDataRow("ExpiredDate") = CDate(dr(8).ToString.Trim)
                End If

                If Len(FixNull(dr(6))) > 50 Then
                    myDataRow("ErrorMsg") = "Error: Length of Manufacturer shouldn't exceed 50"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If
                If Len(FixNull(dr(7))) > 50 Then
                    myDataRow("ErrorMsg") = "Error: Length of ManufacturerPN shouldn't exceed 50"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If
                If Len(FixNull(dr(9))) > 20 Then
                    myDataRow("ErrorMsg") = "Warning: Length of DateCode shouldn't exceed 20."
                End If
                If Len(FixNull(dr(10))) > 20 Then
                    myDataRow("ErrorMsg") = myDataRow("ErrorMsg") & "Warning: Length of LotNo shouldn't exceed 20."
                End If
                If Len(FixNull(dr(11))) > 50 Then
                    myDataRow("ErrorMsg") = "Error: Length of BPCSPN shouldn't exceed 50"
                    MRListData.Tables("MRData").Rows.Add(myDataRow)
                    Continue For
                End If

                If IsDate(myDataRow("DateCode")) Then
                    myDataRow("DateCode") = Mid(myDataRow("DateCode"), 1, 10)
                End If

                myDataRow("UnitQty") = CDec(Format(dr(2), "#.#####"))

                MRListData.Tables("MRData").Rows.Add(myDataRow)
            Next


            Return MRListData

        Catch ex As Exception
            ErrorLogging("InvMigration-CheckFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function CheckBatchFormat(ByVal InvID As String, ByVal BatchListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim i As Integer
        Dim DR_Blank() As DataRow = Nothing

        Try
            'deal with blank rows
            DR_Blank = ExcelData.Tables(0).Select("(Material is null or Material = '') and (unitqty is null or unitqty = '') and (noofpackage is null or noofpackage = '') and (SLOC is null or SLOC = '')")
            If DR_Blank.Length > 0 Then
                For Each drblank As DataRow In DR_Blank
                    ExcelData.Tables(0).Rows.Remove(drblank)
                Next
            End If

            For Each dr As DataRow In ExcelData.Tables(0).Rows
                'deal with blank rows
                'If FixNull(dr(0)) = "" Then
                '    ExcelData.Tables(0).Rows.Remove(dr)
                '    Continue For
                'End If

                RowIndex = RowIndex + 1
                myDataRow = BatchListData.Tables("BatchData").NewRow()
                'myDataRow("Select") = dr(0).ToString.Trim
                myDataRow("InvID") = InvID
                myDataRow("Line") = RowIndex
                myDataRow("Item") = UCase(IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim))
                'myDataRow("Rev") = UCase(IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim))
                'myDataRow("UnitQty") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                'myDataRow("UnitQty") = IIf(dr(2) Is DBNull.Value, 0, FormatNumber(dr(2), 3))
                If FixNull(dr(1)) = "" OrElse FixNull(dr(1)) = 0 Then
                    myDataRow("UnitQty") = 0
                Else
                    myDataRow("UnitQty") = CDec(Format(dr(1), "#.######"))  'FormatNumber(dr(2), 6)
                End If
                myDataRow("NoofPackage") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                myDataRow("SLOC") = UCase(IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim))
                myDataRow("Subinventory") = UCase(IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim))
                myDataRow("Locator") = UCase(IIf(dr(5) Is DBNull.Value, "", dr(5).ToString.Trim))
                'If myDataRow("Locator") <> "" Then
                '    myDataRow("Locator") = Replace(myDataRow("Locator"), "…", "...")
                'End If
                myDataRow("Manufacturer") = SQLString(UCase(IIf(dr(6) Is DBNull.Value, "", CStr(FixNull(dr(6))).Trim)))
                myDataRow("ManufacturerPN") = SQLString(UCase(IIf(dr(7) Is DBNull.Value, "", CStr(FixNull(dr(7))).Trim)))
                myDataRow("ExpiredDate") = IIf(dr(8) Is DBNull.Value, "", dr(8).ToString.Trim)
                'myDataRow("ExpiredDate") = IIf(Fix(dr(8)) = "", "", Date.TryParse(dr(8), dt).ToString.Trim)
                myDataRow("DateCode") = IIf(dr(9) Is DBNull.Value, "", CStr(FixNull(dr(9))).Trim)
                myDataRow("LotNo") = IIf(dr(10) Is DBNull.Value, "", CStr(FixNull(dr(10))).Trim)

                'myDataRow("BPCSPN") = UCase(IIf(dr(11) Is DBNull.Value, "", CStr(FixNull(dr(11))).Trim))
                'If Not dr(10) Is DBNull.Value Then myDataRow("ManufactureDate") = CDate(dr(10))
                If FixNull(dr(0)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Material must be entered"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(1)) = "" Then
                    myDataRow("ErrorMsg") = "Error: UnitQty must be entered"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(1)) = "" AndAlso dr(1) = 0 Then
                    myDataRow("ErrorMsg") = "Error: UnitQty can't be Zero"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(1)) = "" AndAlso dr(1) < 0.00001 Then
                    myDataRow("ErrorMsg") = "Error: UnitQty can't be less than 0.00001"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(2)) = "" Then
                    myDataRow("ErrorMsg") = "Error: NoofPackage must be entered"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(3)) = "" Then
                    myDataRow("ErrorMsg") = "Error: SLOC must be entered"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If

                If Double.TryParse(dr(1), d) = False OrElse dr(1) <= 0 Then
                    myDataRow("ErrorMsg") = "Error: Enter UnitQty only with Positive number"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If

                If dr(2) <= 0 Then
                    myDataRow("ErrorMsg") = "Error: Enter NoOfPackage only with Positive Integer"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                ElseIf dr(2) > 0 And dr(2) <> Int(dr(2)) Then
                    myDataRow("ErrorMsg") = "Error: Enter NoOfPackage only with Integer"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If

                'If FixNull(dr(8)) <> "" AndAlso Date.TryParse(dr(8), dt) = False Then
                '    myDataRow("ErrorMsg") = "Error: Enter ExpiredDate with format: MM/DD/YYYY, do Exceed Today"
                '    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                '    Continue For
                'ElseIf FixNull(dr(8)) <> "" AndAlso Date.TryParse(dr(8), dt) = True Then
                '    Dim datenow As Date
                '    Dim DifOfDate As Integer
                '    datenow = GetDate(OracleERPLogin)
                '    DifOfDate = DateDiff("d", datenow, dr(8))
                '    If DifOfDate < 1 Or Year(dr(8)) > Year(datenow) + 5 Then
                '        myDataRow("ErrorMsg") = "Error: Exp Date " & dr(8) & " is not in correct interval"
                '        BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                '        Continue For
                '    End If
                'End If
                If FixNull(dr(8)) <> "" AndAlso Date.TryParse(dr(8), dt) = False Then
                    myDataRow("ErrorMsg") = "Error: Enter ExpiredDate with format: MM/DD/YYYY."
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If

                If FixNull(dr(8)) <> "" Then
                    myDataRow("ExpiredDate") = CDate(dr(8).ToString.Trim)
                End If

                If Len(FixNull(dr(6))) > 50 Then
                    myDataRow("ErrorMsg") = "Error: Length of Manufacturer shouldn't exceed 50"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If
                If Len(FixNull(dr(7))) > 50 Then
                    myDataRow("ErrorMsg") = "Error: Length of ManufacturerPN shouldn't exceed 50"
                    BatchListData.Tables("BatchData").Rows.Add(myDataRow)
                    Continue For
                End If
                If Len(FixNull(dr(9))) > 20 Then
                    myDataRow("ErrorMsg") = "Warning: Length of DateCode shouldn't exceed 20."
                End If
                If Len(FixNull(dr(10))) > 20 Then
                    myDataRow("ErrorMsg") = myDataRow("ErrorMsg") & "Warning: Length of LotNo shouldn't exceed 20."
                End If
                'If Len(FixNull(dr(11))) > 50 Then
                '    myDataRow("ErrorMsg") = "Error: Length of BPCSPN shouldn't exceed 50"
                '    MRListData.Tables("MRData").Rows.Add(myDataRow)
                '    Continue For
                'End If

                If IsDate(myDataRow("DateCode")) Then
                    myDataRow("DateCode") = Mid(myDataRow("DateCode"), 1, 10)
                End If

                myDataRow("UnitQty") = CDec(Format(dr(1), "#.#####"))

                BatchListData.Tables("BatchData").Rows.Add(myDataRow)
            Next

            Return BatchListData

        Catch ex As Exception
            ErrorLogging("ZSMigration-CheckBatchFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function CheckPIFormat(ByVal InvID As String, ByVal PIListData As DataSet, ByVal ExcelData As DataSet, ByVal Type As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim i As Integer
        Dim DR_Blank() As DataRow = Nothing

        Try
            'deal with blank rows
            DR_Blank = ExcelData.Tables(0).Select("(item is null or item = '') and (qty is null or qty = '') and (subinventory is null or subinventory = '')")
            If DR_Blank.Length > 0 Then
                For Each drblank As DataRow In DR_Blank
                    ExcelData.Tables(0).Rows.Remove(drblank)
                Next
            End If

            For Each dr As DataRow In ExcelData.Tables(0).Rows
                'deal with blank rows
                'If FixNull(dr(0)) = "" Then
                '    ExcelData.Tables(0).Rows.Remove(dr)
                '    Continue For
                'End If

                RowIndex = RowIndex + 1
                myDataRow = PIListData.Tables("PIData").NewRow()
                myDataRow("Item") = UCase(IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim))
                myDataRow("Revision") = UCase(IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim))
                If FixNull(dr(2)) = "" OrElse FixNull(dr(2)) = 0 Then
                    myDataRow("Qty") = 0
                Else
                    myDataRow("Qty") = CDec(Format(dr(2), "#.######"))  'FormatNumber(dr(2), 6)
                End If
                myDataRow("Subinventory") = UCase(IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim))
                myDataRow("Locator") = UCase(IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim))
                If myDataRow("Locator") <> "" Then
                    myDataRow("Locator") = Replace(myDataRow("Locator"), "…", "...")
                End If
                myDataRow("RTLot") = IIf(dr(5) Is DBNull.Value, "", CStr(FixNull(dr(5))).Trim)

                'If Not dr(10) Is DBNull.Value Then myDataRow("ManufactureDate") = CDate(dr(10))
                If FixNull(dr(0)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Item must be entered"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(2)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Qty must be entered"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not Microsoft.VisualBasic.IsNumeric(FixNull(dr(2))) = True Then
                    myDataRow("ErrorMsg") = "Error: Qty must be numeric"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(2)) = "" AndAlso dr(2) = 0 Then
                    myDataRow("ErrorMsg") = "Error: Qty can't be Zero"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                ElseIf Not FixNull(dr(2)) = "" AndAlso System.Math.Abs(dr(2)) < 0.00001 Then
                    myDataRow("ErrorMsg") = "Error: Qty can't be less than 0.00001"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                ElseIf FixNull(dr(3)) = "" Then
                    myDataRow("ErrorMsg") = "Error: Subinventory must be entered"
                    PIListData.Tables("PIData").Rows.Add(myDataRow)
                    Continue For
                End If

                myDataRow("Qty") = CDec(Format(dr(2), "#.#####"))

                PIListData.Tables("PIData").Rows.Add(myDataRow)
            Next

            Return PIListData

        Catch ex As Exception
            ErrorLogging("PIAdjustment-CheckPIFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function ValidMRData(ByVal MRListData As DataSet, ByVal ItemListData As DataSet, ByVal dsAML As DataSet, ByVal SubinvLoc As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Dim i, j, m, n As Integer
            Dim Flag_Rev, Flag_Sub As Boolean
            Dim MaterialNo, Manufacturer, ManufacturerPN As String
            Dim RevList(), RstSubList() As String
            Dim DR() As DataRow = Nothing
            Dim DR_Subinv() As DataRow
            Dim DR_SubinvLoc() As DataRow
            Dim DR_MFR() As DataRow
            Dim DR_MFR_MPN() As DataRow
            Dim DR_AML() As DataRow
            Dim DR_Item() As DataRow
            Dim PostDataRow As Data.DataRow
            Dim CLIDDataRow As Data.DataRow
            Dim ProdDataRow As Data.DataRow
            Dim s_date As String
            Dim ExpDate As String
            Dim itemtype As String
            Dim nolife_flag As Boolean
            Dim LotControl As String
            Dim LifeControl As String
            Dim LifeDays As Integer
            'Dim DateNow As Date

            's_date = Format(DateTime.Now, "MMddHHmmss")
            'TransactionID = CInt(s_date)
            'DateNow = GetDate(OracleLoginData)
            'TransactionID = MRListData.Tables("MRData").Rows(i)("InvID")

            For i = 0 To MRListData.Tables("MRData").Rows.Count - 1
                DR = Nothing
                DR_Subinv = Nothing
                DR_SubinvLoc = Nothing
                DR_MFR = Nothing
                DR_MFR_MPN = Nothing
                DR_AML = Nothing
                DR_Item = Nothing
                itemtype = ""
                nolife_flag = False

                Flag_Rev = False
                Flag_Sub = False
                MaterialNo = MRListData.Tables("MRData").Rows(i)("Item")
                DR = ItemListData.Tables(0).Select(" p_item_num = '" & MaterialNo & "'and o_valid_flag = 'Y'")
                If Not DR.Length > 0 Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Item is invalid"
                    Continue For
                End If
                If (FixNull(MRListData.Tables("MRData").Rows(i)("Revision")) <> "" And FixNull(DR(0)("o_revision_control_code")) = 1) Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Item with no Revision Control. Revision is not required"
                    Continue For
                End If

                If (FixNull(MRListData.Tables("MRData").Rows(i)("Revision")) = "" And FixNull(DR(0)("o_revision_control_code")) = 2) Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Item with Revision Control. Revision is required"
                    Continue For
                End If

                If FixNull(DR(0)("o_revision_control_code")) = 2 And FixNull(MRListData.Tables("MRData").Rows(i)("Revision")) <> "" Then
                    Erase RevList
                    If InStr(FixNull(DR(0)("o_revlist")), ",") > 0 Then
                        RevList = Split(FixNull(DR(0)("o_revlist")), ",")
                        For j = LBound(RevList) To UBound(RevList)
                            If RevList(j) = MRListData.Tables("MRData").Rows(i)("Revision") Then  'This is very important, txtRev place at first!
                                Flag_Rev = True
                                Exit For
                            End If
                        Next
                    Else
                        If FixNull(MRListData.Tables("MRData").Rows(i)("Revision")) = FixNull(DR(0)("o_revlist")) Then
                            Flag_Rev = True
                        End If
                    End If
                    If Flag_Rev = False Then
                        MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Item with Revision Control, but Revision is invalid"
                        Continue For
                    End If
                End If

                ''''''not checking subinv/locator for U-Turn ''''''

                'DR_Subinv = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & MRListData.Tables("MRData").Rows(i)("Subinventory") & "'")
                'If Not DR_Subinv.Length > 0 Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Subinventory is invalid"
                '    Continue For
                'End If

                'Erase RstSubList
                'If InStr(FixNull(DR(0)("o_sublist")), ",") > 0 Then
                '    RstSubList = Split(FixNull(DR(0)("o_sublist")), ",")
                '    For j = LBound(RstSubList) To UBound(RstSubList)
                '        If RstSubList(j) = MRListData.Tables("MRData").Rows(i)("Subinventory") Then  'This is very important, txtRev place at first!
                '            Flag_Sub = True
                '            Exit For
                '        End If
                '    Next
                'Else
                '    If FixNull(MRListData.Tables("MRData").Rows(i)("Subinventory")) = FixNull(DR(0)("o_sublist")) Then
                '        Flag_Sub = True
                '    End If
                '    'if no restrict list, pass the subinventory
                '    If FixNull(DR(0)("o_sublist")) = "" Then
                '        Flag_Sub = True
                '    End If
                'End If
                'If Flag_Sub = False Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Subinventory is not in the restricted list"
                '    Continue For
                'End If

                'DR_SubinvLoc = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & MRListData.Tables("MRData").Rows(i)("Subinventory") & "'and LOCATOR = '" & FixNull(MRListData.Tables("MRData").Rows(i)("Locator")) & "'")
                ''DR_SubinvLoc = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & MRListData.Tables("MRData").Rows(i)("Subinventory") & "'and LOCATOR = '" & "" & "'")

                'If Not DR_SubinvLoc.Length > 0 Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Locator is invalid"
                '    Continue For
                'End If

                LotControl = ""
                LifeControl = ""
                LifeDays = 0
                If FixNull(DR(0)("o_lot_control_code")) = "2" Then
                    LotControl = "Y"
                End If

                If LotControl = "Y" Then
                    'o_shelf_life_ctrl: 1 no control ,2 shelf life days ,4 user defined
                    If FixNull(DR(0)("o_shelf_life_ctrl")) <> "" Then
                        If FixNull(DR(0)("o_shelf_life_ctrl")) = "1" Then
                            LifeControl = "N"
                        ElseIf FixNull(DR(0)("o_shelf_life_ctrl")) = "2" OrElse FixNull(DR(0)("o_shelf_life_ctrl")) = "4" Then
                            If DR(0)("o_shelf_life_days") Is DBNull.Value Then
                                MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Shelf Life Days setting invalid in Oracle"
                                Continue For
                            ElseIf Not DR(0)("o_shelf_life_days") Is DBNull.Value Then
                                LifeDays = DR(0)("o_shelf_life_days")
                                LifeControl = "Y"
                            End If
                        End If
                    End If
                End If

                If LifeControl = "Y" AndAlso LifeDays > 0 AndAlso FixNull(MRListData.Tables("MRData").Rows(i)("ExpiredDate")) = "" Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: ExpiredDate is required"
                    Continue For
                End If
                If LifeControl <> "Y" AndAlso FixNull(MRListData.Tables("MRData").Rows(i)("ExpiredDate")) <> "" Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Warning: No shelf life control, entered ExpiredDate won't be considered."
                End If

                'If (FixNull(DR(0)("shelf_life_code")) = "2" OrElse FixNull(DR(0)("shelf_life_code")) = "4") AndAlso DR(0)("o_shelf_life_days") Is DBNull.Value Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: Shelf Life Days setting invalid in Oracle"
                '    Continue For
                'End If
                'If (FixNull(DR(0)("shelf_life_code")) = "2" OrElse FixNull(DR(0)("shelf_life_code")) = "4") AndAlso DR(0)("o_shelf_life_days") > 0 And FixNull(MRListData.Tables("MRData").Rows(i)("ExpiredDate")) = "" Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Error: ExpiredDate is required"
                '    Continue For
                'End If


                'No shelf life control, don't consider entered Expired Date, just add 3 years as expired date
                If LifeControl = "Y" AndAlso Not LifeDays > 0 Then
                    nolife_flag = True
                    If FixNull(MRListData.Tables("MRData").Rows(i)("ExpiredDate")) <> "" Then
                        MRListData.Tables("MRData").Rows(i)("ErrorMsg") = "Warning: No shelf life days setting, entered ExpiredDate won't be considered."
                    End If
                End If

                Manufacturer = FixNull(MRListData.Tables("MRData").Rows(i)("Manufacturer"))
                If Manufacturer <> "" Then
                    Manufacturer = SQLString(Manufacturer)
                End If
                ManufacturerPN = FixNull(MRListData.Tables("MRData").Rows(i)("ManufacturerPN"))
                If ManufacturerPN <> "" Then
                    ManufacturerPN = SQLString(ManufacturerPN)
                End If
                'Only check AML for raw material
                'If DR(0)("o_type_name") = "RM" Then
                'If dsAML.Tables("AMLData").Rows.Count > 0 Then
                DR_MFR_MPN = dsAML.Tables("AMLData").Select(" MaterialNo = '" & MaterialNo & "' and MFR = '" & Manufacturer & "' and MPN = '" & ManufacturerPN & "'")
                If Not DR_MFR_MPN.Length > 0 Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = MRListData.Tables("MRData").Rows(i)("ErrorMsg") & "Warning: MFR & MPN not match iProd information."
                End If
                'End If

                'DR_MFR_MPN = dsAML.Tables("AMLData").Select(" MaterialNo = '" & MaterialNo & "' and MFR = '" & Manufacturer & "' and MPN = '" & ManufacturerPN & "'")
                'If Not DR_MFR_MPN.Length > 0 Then
                '    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = MRListData.Tables("MRData").Rows(i)("ErrorMsg") & "Warning: MFR & MPN not match iProd information."
                'End If
                'End If

                If Len(Trim(MRListData.Tables("MRData").Rows(i)("DateCode"))) > 20 Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = MRListData.Tables("MRData").Rows(i)("ErrorMsg") & "Warning: Length of DateCode shouldn't exceed 20."
                End If
                If Len(Trim(MRListData.Tables("MRData").Rows(i)("LotNo"))) > 20 Then
                    MRListData.Tables("MRData").Rows(i)("ErrorMsg") = MRListData.Tables("MRData").Rows(i)("ErrorMsg") & "Warning: Length of LotNo shouldn't exceed 20."
                End If

                'Add row for PostData
                Dim id As String
                Dim TransactionID As Double
                id = MRListData.Tables("MRData").Rows(i)("InvID")
                id = OracleLoginData.OrgCode & id
                TransactionID = CDbl(id)

                PostDataRow = MRListData.Tables("accountreceipt_table").NewRow()
                PostDataRow("p_timeout") = 1800000
                PostDataRow("p_organization_code") = OracleLoginData.OrgCode                       'p_organization_code
                'PostDataRow("p_transaction_header_id") = MRListData.Tables("MRData").Rows(i)("InvID")             'p_transaction_header_id
                PostDataRow("p_transaction_header_id") = TransactionID            'p_transaction_header_id

                If DR(0)("o_type_name") = "RM" Or DR(0)("o_type_name") = "SA" Then
                    PostDataRow("p_transaction_source_name") = "INVTY CONV - RM/SA"
                ElseIf DR(0)("o_type_name") = "FG" Then
                    PostDataRow("p_transaction_source_name") = "INVTY CONV - FG"
                End If

                PostDataRow("p_transaction_uom") = DR(0)("o_uom_code")         'p_transaction_uom
                PostDataRow("p_source_line_id") = 99                               'p_source_line_id
                PostDataRow("p_source_header_id") = 99                             'p_source_header_id
                PostDataRow("p_user_id") = OracleLoginData.UserID                   'p_user_id
                PostDataRow("p_item_segment1") = MaterialNo
                PostDataRow("p_item_revision") = FixNull(MRListData.Tables("MRData").Rows(i)("Revision"))    'Null value ???????????????????
                PostDataRow("p_subinventory_destination") = MRListData.Tables("MRData").Rows(i)("Subinventory")    'p_subinventory_destination  destination
                PostDataRow("p_locator_destination") = FixNull(MRListData.Tables("MRData").Rows(i)("Locator"))

                If FixNull(DR(0)("o_lot_control_code")) = 2 Then
                    ExpDate = ""
                    If FixNull(MRListData.Tables("MRData").Rows(i)("ExpiredDate")) <> "" AndAlso LifeControl = "Y" AndAlso LifeDays > 0 Then
                        PostDataRow("p_lot_expiration_date") = MRListData.Tables("MRData").Rows(i)("ExpiredDate")                 'p_lot_expiration_date
                        ExpDate = Format(CDate(PostDataRow("p_lot_expiration_date")), "yyyyMMdd")
                    ElseIf LifeControl = "Y" AndAlso LifeDays = 0 Then
                        'myDataRow("p_lot_expiration_date") = DateTime.Now.AddYears(3)
                        PostDataRow("p_lot_expiration_date") = DateTime.Now.AddYears(3)
                        ExpDate = Format(CDate(PostDataRow("p_lot_expiration_date")), "yyyyMMdd")
                    End If
                    PostDataRow("p_lot_number") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")
                Else

                End If

                PostDataRow("p_reason_code") = "ST-SIB"                                                                   'To be decided
                PostDataRow("p_transaction_reference") = Mid(MRListData.Tables("MRData").Rows(i)("BPCSPN").ToString.Trim, 1, 50)      'To be decided
                PostDataRow("p_transaction_quantity") = MRListData.Tables("MRData").Rows(i)("UnitQty") * MRListData.Tables("MRData").Rows(i)("NoofPackage")  'p_transaction_quantity
                PostDataRow("p_primary_quantity") = MRListData.Tables("MRData").Rows(i)("UnitQty") * MRListData.Tables("MRData").Rows(i)("NoofPackage")                     'p_primary_quantity
                PostDataRow("o_return_status") = ""                                  'o_return_status
                PostDataRow("o_return_message") = ""                                 'o_return_message

                'For Stock Take - Begin
                PostDataRow("p_transaction_source_name") = "MAT RCPT (PRDN RET)"
                PostDataRow("p_reason_code") = "CC-CCA_SM"
                'For Stock Take - End

                'For item issue correction - Begin
                PostDataRow("p_transaction_source_name") = "LOT CONTROL INV RCPT"
                PostDataRow("p_reason_code") = "MR-LOT"
                'For item issue correction -End

                MRListData.Tables("accountreceipt_table").Rows.Add(PostDataRow)
                If DR(0)("o_type_name") = "RM" Then
                    For m = 0 To MRListData.Tables("MRData").Rows(i)("NoofPackage") - 1
                        'Add row for CLIDData
                        'Dim Exp_Date As String
                        Dim RecDate As String
                        ' Get Traceability Level and SafeCode according to Material Group              'Add by Yudy 11/28/2008
                        'Dim SafeCode As String
                        'Dim TraceLevel As MGTraceData
                        '            LabelForMiscReceipt.TraceLevel = "PT"    ' "NT" not needed   "CT" must needed  "PT" optional
                        'SafeCode = ""

                        RecDate = Date.Now

                        'Exp_Date = MRListData.Tables("MRData").Rows(i)("ExpiredDate")
                        'If Exp_Date Is Nothing Then
                        '    Exp_Date = Replace(DateTime.Now.AddYears(3), "-", "/")
                        'End If

                        DR_AML = dsAML.Tables("AMLData").Select(" MaterialNo = '" & MaterialNo & "' ")
                        DR_Item = dsAML.Tables("ItemData").Select(" MaterialNo = '" & MaterialNo & "' ")

                        CLIDDataRow = MRListData.Tables("CLIDData").NewRow()
                        CLIDDataRow("InvID") = MRListData.Tables("MRData").Rows(i)("InvID")
                        CLIDDataRow("Line") = MRListData.Tables("MRData").Rows(i)("Line")
                        CLIDDataRow("NewCLID") = ""
                        CLIDDataRow("OrgCode") = OracleLoginData.OrgCode
                        CLIDDataRow("MaterialNo") = MaterialNo
                        CLIDDataRow("MaterialRevision") = MRListData.Tables("MRData").Rows(i)("Revision")
                        CLIDDataRow("Qty") = MRListData.Tables("MRData").Rows(i)("UnitQty")
                        CLIDDataRow("UOM") = DR(0)("o_uom_code")
                        CLIDDataRow("QtyBaseUOM") = MRListData.Tables("MRData").Rows(i)("UnitQty")
                        CLIDDataRow("BaseUOM") = DR(0)("o_uom_code")
                        CLIDDataRow("DateCode") = SQLString(Mid(FixNull(MRListData.Tables("MRData").Rows(i)("DateCode")), 1, 20))
                        CLIDDataRow("LotNo") = SQLString(Mid(FixNull(MRListData.Tables("MRData").Rows(i)("LotNo")), 1, 20))
                        CLIDDataRow("COO") = DBNull.Value
                        CLIDDataRow("RecDocNo") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")   'MRListData.Tables("MRData").Rows(i)("InvID")
                        CLIDDataRow("RecDocItem") = DBNull.Value
                        CLIDDataRow("CreatedBy") = OracleLoginData.User.ToUpper
                        CLIDDataRow("Printed") = "True"
                        CLIDDataRow("VendorID") = ""
                        CLIDDataRow("RecDate") = CDate(RecDate)
                        If DR_Item.Length > 0 Then
                            CLIDDataRow("RoHS") = DR_Item(0)("RoHS")
                        End If
                        CLIDDataRow("PurOrdNo") = ""
                        CLIDDataRow("PurOrdItem") = ""
                        CLIDDataRow("InvoiceNo") = ""
                        CLIDDataRow("BillofLading") = ""
                        CLIDDataRow("DN") = ""
                        CLIDDataRow("HeaderText") = ""
                        'If FixNull(Exp_Date) = "" Then
                        '    CLIDDataRow("ExpDate") = DateTime.Now.AddYears(3)   'DBNull.Value
                        'Else
                        '    CLIDDataRow("ExpDate") = CDate(Exp_Date)
                        'End If
                        'CLIDDataRow("ExpDate") = CDate(Exp_Date)
                        'If FixNull(DR(0)("o_lot_control_code")) = 2 Then
                        '    CLIDDataRow("RTLot") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")  'MRListData.Tables("MRData").Rows(i)("InvID")
                        '    CLIDDataRow("ExpDate") = PostDataRow("p_lot_expiration_date")
                        'Else
                        '    CLIDDataRow("RTLot") = DBNull.Value
                        '    CLIDDataRow("ExpDate") = DBNull.Value
                        'End If
                        If LotControl = "Y" Then
                            CLIDDataRow("RTLot") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")
                        Else
                            CLIDDataRow("RTLot") = DBNull.Value
                        End If
                        If LifeControl = "Y" Then
                            CLIDDataRow("ExpDate") = PostDataRow("p_lot_expiration_date")
                        Else
                            CLIDDataRow("ExpDate") = DBNull.Value
                        End If

                        CLIDDataRow("ProdDate") = DBNull.Value
                        CLIDDataRow("ReasonCode") = ""
                        If DR(0)("o_routing_id").ToString = "2" Then
                            CLIDDataRow("StockType") = "QP"
                        Else
                            CLIDDataRow("StockType") = "FTS"
                        End If
                        CLIDDataRow("MaterialDesc") = SQLString(DR(0)("o_item_desc"))
                        CLIDDataRow("VendorName") = ""
                        CLIDDataRow("VendorPN") = ""
                        CLIDDataRow("SLOC") = MRListData.Tables("MRData").Rows(i)("Subinventory")
                        CLIDDataRow("StorageBin") = MRListData.Tables("MRData").Rows(i)("Locator")
                        CLIDDataRow("Operator") = OracleLoginData.User.ToUpper
                        CLIDDataRow("IsTraceable") = "PT"
                        CLIDDataRow("MatDocNo") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")
                        CLIDDataRow("MatDocItem") = DBNull.Value
                        If InStr(Manufacturer, "^") > 0 Then
                            CLIDDataRow("Manufacturer") = SQLString(Replace(Manufacturer, "^", "~"))
                        Else
                            CLIDDataRow("Manufacturer") = SQLString(Manufacturer)
                        End If
                        If InStr(ManufacturerPN, "^") > 0 Then
                            CLIDDataRow("ManufacturerPN") = SQLString(Replace(ManufacturerPN, "^", "~"))
                        Else
                            CLIDDataRow("ManufacturerPN") = SQLString(ManufacturerPN)
                        End If
                        'CLIDDataRow("Manufacturer") = Manufacturer
                        'CLIDDataRow("ManufacturerPN") = ManufacturerPN
                        If DR_AML.Length > 0 Then
                            CLIDDataRow("QMLStatus") = SQLString(FixNull(DR_AML(0)("AMLStatus")))
                        End If
                        If DR_Item.Length > 0 Then
                            CLIDDataRow("AddlData") = SQLString(FixNull(DR_Item(0)("AddlData")))
                            CLIDDataRow("Stemp") = SQLString(FixNull(DR_Item(0)("Stemp")))
                            If InStr(MRListData.Tables("MRData").Rows(i)("BPCSPN"), "^") > 0 Then
                                CLIDDataRow("MSL") = SQLString(Mid(FixNull(DR_Item(0)("MSL")) & " (" & Replace(FixNull(MRListData.Tables("MRData").Rows(i)("BPCSPN")), "^", "~") & ")", 1, 50))
                            Else
                                CLIDDataRow("MSL") = SQLString(Mid(FixNull(DR_Item(0)("MSL")) & " (" & FixNull(MRListData.Tables("MRData").Rows(i)("BPCSPN")) & ")", 1, 50))
                            End If
                            'CLIDDataRow("MSL") = DR_Item(0)("MSL") & " (" & MRListData.Tables("MRData").Rows(i)("BPCSPN") & ")"
                        Else
                            CLIDDataRow("AddlData") = ""
                            CLIDDataRow("Stemp") = ""
                            If InStr(MRListData.Tables("MRData").Rows(i)("BPCSPN"), "^") > 0 Then
                                CLIDDataRow("MSL") = SQLString(Mid(" (" & Replace(FixNull(MRListData.Tables("MRData").Rows(i)("BPCSPN")), "^", "~") & ")", 1, 50))
                            Else
                                CLIDDataRow("MSL") = SQLString(Mid(" (" & FixNull(MRListData.Tables("MRData").Rows(i)("BPCSPN")) & ")", 1, 50))
                            End If
                            'CLIDDataRow("MSL") = DR_Item(0)("MSL") & " (" & MRListData.Tables("MRData").Rows(i)("BPCSPN") & ")"
                        End If
                        MRListData.Tables("CLIDData").Rows.Add(CLIDDataRow)
                    Next
                ElseIf DR(0)("o_type_name") = "FG" Or DR(0)("o_type_name") = "SA" Then
                    For n = 0 To MRListData.Tables("MRData").Rows(i)("NoofPackage") - 1
                        'Add row for ProdData
                        ProdDataRow = MRListData.Tables("ProdData").NewRow()
                        ProdDataRow("InvID") = MRListData.Tables("MRData").Rows(i)("InvID")
                        ProdDataRow("Line") = MRListData.Tables("MRData").Rows(i)("Line")
                        ProdDataRow("CLID") = ""
                        ProdDataRow("PurOrdNo") = Mid(MRListData.Tables("MRData").Rows(i)("InvID") & "(" & MRListData.Tables("MRData").Rows(i)("BPCSPN") & ")", 1, 50)
                        ProdDataRow("OrgCode") = OracleLoginData.OrgCode
                        ProdDataRow("MaterialNo") = MaterialNo
                        ProdDataRow("MaterialDesc") = SQLString(DR(0)("o_item_desc"))
                        ProdDataRow("MaterialRevision") = MRListData.Tables("MRData").Rows(i)("Revision")
                        ProdDataRow("QtyBaseUOM") = MRListData.Tables("MRData").Rows(i)("UnitQty")
                        ProdDataRow("BaseUOM") = DR(0).Item(6)
                        'If FixNull(DR(0)("o_lot_control_code")) = 2 Then
                        '    ProdDataRow("RTLot") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")  'MRListData.Tables("MRData").Rows(i)("InvID")
                        '    ProdDataRow("ExpDate") = PostDataRow("p_lot_expiration_date")
                        'Else
                        '    ProdDataRow("RTLot") = DBNull.Value
                        '    ProdDataRow("ExpDate") = DBNull.Value
                        'End If
                        If LotControl = "Y" Then
                            ProdDataRow("RTLot") = ExpDate & MRListData.Tables("MRData").Rows(i)("InvID")
                        Else
                            ProdDataRow("RTLot") = DBNull.Value
                        End If
                        If LifeControl = "Y" Then
                            ProdDataRow("ExpDate") = PostDataRow("p_lot_expiration_date")
                        Else
                            ProdDataRow("ExpDate") = DBNull.Value
                        End If
                        ProdDataRow("DateCode") = SQLString(Mid(FixNull(MRListData.Tables("MRData").Rows(i)("DateCode")), 1, 20))
                        ProdDataRow("LotNo") = SQLString(Mid(FixNull(MRListData.Tables("MRData").Rows(i)("LotNo")), 1, 20))
                        ProdDataRow("SubInv") = MRListData.Tables("MRData").Rows(i)("Subinventory")
                        ProdDataRow("Locator") = MRListData.Tables("MRData").Rows(i)("Locator")
                        MRListData.Tables("ProdData").Rows.Add(ProdDataRow)
                    Next
                End If
            Next
            Return MRListData

        Catch ex As Exception
            ErrorLogging("InvMigration-ValidMRData", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function ValidPIData(ByVal PIListData As DataSet, ByVal ItemListData As DataSet, ByVal ItemOnhand As DataSet, ByVal SubinvLoc As DataSet, ByVal Type As String, ByVal TransactionID As Long, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Dim i, j, m, n As Integer
            Dim Flag_Rev, Flag_Sub As Boolean
            Dim MaterialNo, Manufacturer, ManufacturerPN As String
            Dim RevList(), RstSubList() As String
            Dim DR() As DataRow = Nothing
            Dim DR_Subinv() As DataRow
            Dim DR_SubinvLoc() As DataRow
            Dim DR_MFR() As DataRow
            Dim DR_MFR_MPN() As DataRow
            Dim DR_AML() As DataRow
            Dim DR_Item() As DataRow
            Dim DR_Onhand() As DataRow
            Dim PostDataRow As Data.DataRow
            Dim CLIDDataRow As Data.DataRow
            Dim ProdDataRow As Data.DataRow
            Dim s_date As String
            Dim ExpDate As String
            Dim itemtype As String
            Dim nolife_flag As Boolean
            Dim LotControl As String
            Dim LifeControl As String
            Dim LifeDays As Integer
            'Dim DateNow As Date

            's_date = Format(DateTime.Now, "MMddHHmmss")
            'TransactionID = CInt(s_date)
            'DateNow = GetDate(OracleLoginData)
            'TransactionID = MRListData.Tables("MRData").Rows(i)("InvID")
            j = 0
            For i = 0 To PIListData.Tables("PIData").Rows.Count - 1
                DR = Nothing
                DR_Subinv = Nothing
                DR_SubinvLoc = Nothing
                DR_MFR = Nothing
                DR_MFR_MPN = Nothing
                DR_AML = Nothing
                DR_Item = Nothing
                DR_Onhand = Nothing
                itemtype = ""
                nolife_flag = False

                Flag_Rev = False
                Flag_Sub = False
                MaterialNo = PIListData.Tables("PIData").Rows(i)("Item")
                DR = ItemListData.Tables(0).Select(" p_item_num = '" & MaterialNo & "'and o_valid_flag = 'Y'")
                If Not DR.Length > 0 Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Item is invalid"
                    Continue For
                End If
                If (FixNull(PIListData.Tables("PIData").Rows(i)("Revision")) <> "" And FixNull(DR(0)("o_revision_control_code")) = 1) Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Item with no Revision Control. Revision is not required"
                    Continue For
                End If

                If (FixNull(PIListData.Tables("PIData").Rows(i)("Revision")) = "" And FixNull(DR(0)("o_revision_control_code")) = 2) Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Item with Revision Control. Revision is required"
                    Continue For
                End If

                If FixNull(DR(0)("o_revision_control_code")) = 2 And FixNull(PIListData.Tables("PIData").Rows(i)("Revision")) <> "" Then
                    Erase RevList
                    If InStr(FixNull(DR(0)("o_revlist")), ",") > 0 Then
                        RevList = Split(FixNull(DR(0)("o_revlist")), ",")
                        For j = LBound(RevList) To UBound(RevList)
                            If RevList(j) = PIListData.Tables("PIData").Rows(i)("Revision") Then  'This is very important, txtRev place at first!
                                Flag_Rev = True
                                Exit For
                            End If
                        Next
                    Else
                        If FixNull(PIListData.Tables("PIData").Rows(i)("Revision")) = FixNull(DR(0)("o_revlist")) Then
                            Flag_Rev = True
                        End If
                    End If
                    If Flag_Rev = False Then
                        PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Item with Revision Control, but Revision is invalid"
                        Continue For
                    End If
                End If

                DR_Subinv = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & PIListData.Tables("PIData").Rows(i)("Subinventory") & "'")
                If Not DR_Subinv.Length > 0 Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Subinventory is invalid"
                    Continue For
                End If

                Erase RstSubList
                If InStr(FixNull(DR(0)("o_sublist")), ",") > 0 Then
                    RstSubList = Split(FixNull(DR(0)("o_sublist")), ",")
                    For j = LBound(RstSubList) To UBound(RstSubList)
                        If RstSubList(j) = PIListData.Tables("PIData").Rows(i)("Subinventory") Then  'This is very important, txtRev place at first!
                            Flag_Sub = True
                            Exit For
                        End If
                    Next
                Else
                    If FixNull(PIListData.Tables("PIData").Rows(i)("Subinventory")) = FixNull(DR(0)("o_sublist")) Then
                        Flag_Sub = True
                    End If
                    'if no restrict list, pass the subinventory
                    If FixNull(DR(0)("o_sublist")) = "" Then
                        Flag_Sub = True
                    End If
                End If
                If Flag_Sub = False Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Subinventory is not in the restricted list"
                    Continue For
                End If

                DR_SubinvLoc = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & PIListData.Tables("PIData").Rows(i)("Subinventory") & "'and LOCATOR = '" & FixNull(PIListData.Tables("PIData").Rows(i)("Locator")) & "'")
                'DR_SubinvLoc = SubinvLoc.Tables(0).Select(" SUBINVENTORY = '" & MRListData.Tables("MRData").Rows(i)("Subinventory") & "'and LOCATOR = '" & "" & "'")

                If Not DR_SubinvLoc.Length > 0 Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Locator is invalid"
                    Continue For
                End If

                LotControl = ""
                LifeControl = ""
                LifeDays = 0
                If FixNull(DR(0)("o_lot_control_code")) = "2" Then
                    LotControl = "Y"
                End If

                If LotControl = "Y" AndAlso FixNull(PIListData.Tables("PIData").Rows(i)("RTLot")) = "" Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: RTLot is required"
                    Continue For
                End If

                If Not LotControl = "Y" AndAlso FixNull(PIListData.Tables("PIData").Rows(i)("RTLot")) <> "" Then
                    PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: No Lot Control. Pls remove the RTLot#"
                    Continue For
                End If

                If Type = "ISSUE" Then
                    DR_Onhand = ItemOnhand.Tables(0).Select(" p_item_num = '" & PIListData.Tables("PIData").Rows(i)("Item") & "'and p_item_rev = '" & FixNull(PIListData.Tables("PIData").Rows(i)("Revision")) & "' and p_subinventory = '" & PIListData.Tables("PIData").Rows(i)("Subinventory") & "' and p_locator = '" & FixNull(PIListData.Tables("PIData").Rows(i)("Locator")) & "' and p_lot_number = '" & FixNull(PIListData.Tables("PIData").Rows(i)("RTLot")) & "'")
                    If Not DR_Onhand.Length > 0 Then
                        PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: No onhand accordingly in Oracle"
                        Continue For
                    End If
                    If DR_Onhand(0)("o_available_qty") <= 0 Then
                        PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: no onhand accordingly in Oracle"
                        Continue For
                    ElseIf DR_Onhand(0)("o_available_qty") > 0 AndAlso DR_Onhand(0)("o_available_qty").ToString < PIListData.Tables("PIData").Rows(i)("Qty") * -1 Then
                        PIListData.Tables("PIData").Rows(i)("ErrorMsg") = "Error: Onhand in Oracle " & DR_Onhand(0)("o_available_qty") & " < Issue Qty " & PIListData.Tables("PIData").Rows(i)("Qty") * -1
                        Continue For
                    End If
                End If

                'Add row for PostData
                If Type = "RCPT" Then
                    PostDataRow = PIListData.Tables("accountreceipt_table").NewRow()
                    PostDataRow("p_transaction_source_name") = "PI ADJ (RECEIPT)"
                ElseIf Type = "ISSUE" Then
                    PostDataRow = PIListData.Tables("accountissue_table").NewRow()
                    PostDataRow("p_transaction_source_name") = "PI ADJ (ISSUE)"
                End If
                PostDataRow("p_timeout") = 1800000
                PostDataRow("p_organization_code") = OracleLoginData.OrgCode                       'p_organization_code
                'PostDataRow("p_transaction_header_id") = MRListData.Tables("MRData").Rows(i)("InvID")             'p_transaction_header_id
                PostDataRow("p_transaction_header_id") = TransactionID            'p_transaction_header_id

                'If DR(0)("o_type_name") = "RM" Or DR(0)("o_type_name") = "SA" Then
                '    PostDataRow("p_transaction_source_name") = "INVTY CONV - RM/SA"
                'ElseIf DR(0)("o_type_name") = "FG" Then
                '    PostDataRow("p_transaction_source_name") = "INVTY CONV - FG"
                'End If
                If OracleLoginData.OrgCode = "580" AndAlso DR(0)("o_type_name") = "FG" Then
                    PostDataRow("p_transaction_source_name") = "INVTY CONV - FG"
                ElseIf OracleLoginData.OrgCode = "580" AndAlso DR(0)("o_type_name") <> "FG" Then
                    PostDataRow("p_transaction_source_name") = "INVTY CONV - RM/SA"
                End If

                PostDataRow("p_transaction_uom") = DR(0)("o_uom_code")         'p_transaction_uom
                PostDataRow("p_source_line_id") = 99                               'p_source_line_id
                PostDataRow("p_source_header_id") = 99                             'p_source_header_id
                PostDataRow("p_user_id") = OracleLoginData.UserID                   'p_user_id
                PostDataRow("p_item_segment1") = MaterialNo
                PostDataRow("p_item_revision") = FixNull(PIListData.Tables("PIData").Rows(i)("Revision"))    'Null value ???????????????????
                PostDataRow("p_lot_number") = PIListData.Tables("PIData").Rows(i)("RTLot")
                If Type = "RCPT" AndAlso LotControl = "Y" Then
                    PostDataRow("p_lot_expiration_date") = DateTime.Now.AddYears(3)
                End If
                PostDataRow("p_reason_code") = ""                                                                   'To be decided
                PostDataRow("p_transaction_reference") = ""      'To be decided
                PostDataRow("o_return_status") = ""                                  'o_return_status
                PostDataRow("o_return_message") = ""                                 'o_return_message
                If Type = "RCPT" Then
                    PostDataRow("p_subinventory_destination") = PIListData.Tables("PIData").Rows(i)("Subinventory")    'p_subinventory_destination  destination
                    PostDataRow("p_locator_destination") = FixNull(PIListData.Tables("PIData").Rows(i)("Locator"))
                    PostDataRow("p_transaction_quantity") = PIListData.Tables("PIData").Rows(i)("Qty") 'p_transaction_quantity
                    PostDataRow("p_primary_quantity") = PIListData.Tables("PIData").Rows(i)("Qty")                    'p_primary_quantity
                    PIListData.Tables("accountreceipt_table").Rows.Add(PostDataRow)
                ElseIf Type = "ISSUE" Then
                    'For Stock Take - Begin
                    PostDataRow("p_transaction_source_name") = "MAT ISSUE (PRDN RET)"
                    PostDataRow("p_reason_code") = "CC-CCA_SM"
                    'For Stock Take - End

                    PostDataRow("p_subinventory_source") = PIListData.Tables("PIData").Rows(i)("Subinventory")    'p_subinventory_destination  destination
                    PostDataRow("p_locator_source") = FixNull(PIListData.Tables("PIData").Rows(i)("Locator"))
                    PostDataRow("p_transaction_quantity") = PIListData.Tables("PIData").Rows(i)("Qty") * -1       'p_transaction_quantity
                    PostDataRow("p_primary_quantity") = PIListData.Tables("PIData").Rows(i)("Qty") * -1           'p_primary_quantity
                    PIListData.Tables("accountissue_table").Rows.Add(PostDataRow)
                End If
            Next
            Return PIListData

        Catch ex As Exception
            ErrorLogging("PIAdjustment-ValidPIData", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function ValidINVNo(ByVal INVNo As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr1, Sqlstr2 As String
                Sqlstr1 = String.Format("SELECT InvID,Line,Item,Revision,UnitQty,NoofPackage,Subinventory,Locator,Manufacturer,ManufacturerPN,ExpiredDate,DateCode,LotNo,BPCSPN from T_InvMaster with (nolock) where InvID = '{0}'", INVNo)
                Sqlstr2 = String.Format("SELECT InvID,Line,CLID from T_InvLabel with (nolock) where InvID = '{0}'", INVNo)
                Dim sql() As String = {Sqlstr1, Sqlstr2}
                Dim table() As String = {"MRData", "CLIDTable"}
                Return da.ExecuteDataSet(sql, table)            'dataset里有两个table,tablename 是 "PDTOItems"和"test"
            Catch ex As Exception
                ErrorLogging("InvMigration-ValidINVNo", OracleERPLogin.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
        End Using
    End Function

    Public Function ValidBatchNo(ByVal BatchNo As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr1, Sqlstr2 As String
                Sqlstr1 = String.Format("SELECT InvID,Line,Item,UnitQty,NoofPackage,SLOC,Subinventory,Locator,Manufacturer,ManufacturerPN,ExpiredDate,DateCode,LotNo from T_InvMaster with (nolock) where InvID = '{0}'", BatchNo)
                Sqlstr2 = String.Format("SELECT InvID,Line,CLID from T_InvLabel with (nolock) where InvID = '{0}'", BatchNo)
                Dim sql() As String = {Sqlstr1, Sqlstr2}
                Dim table() As String = {"BatchData", "CLIDTable"}
                Return da.ExecuteDataSet(sql, table)            'datasetÀïÓÐÁ½¸ötable,tablename ÊÇ "PDTOItems"ºÍ"test"
            Catch ex As Exception
                ErrorLogging("ZSMigration-ValidBatchNo", OracleERPLogin.User, ex.Message & ex.Source, "E")
                Throw ex
            End Try
        End Using
    End Function

    Public Function PostMR(ByVal MRListData As DataSet, ByVal MoveType As String, ByVal Printer As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Dim TransactionID As Double
            Dim ID As String
            Dim MiscType As String
            Dim MRListData_Return As New DataSet
            Dim DR() As DataRow = Nothing
            'TransactionID = GetNextInvID(OracleLoginData)
            MiscType = ""
            MRListData_Return = MRListData.Clone

            ID = MRListData.Tables("MRData").Rows(0)("InvID")
            ID = OracleLoginData.OrgCode & ID
            TransactionID = CDbl(ID)


            'MRListData = account_alias_receipt(MRListData, OracleLoginData, MoveType, TransactionID, MiscType)
            MRListData = account_alias_batch_receipt(MRListData, OracleLoginData, MoveType, TransactionID, MiscType)

            DR = MRListData.Tables("accountreceipt_table").Select("o_return_status = 'N'")
            If DR.Length = 0 Then
                'Dim CLIDs As DataSet = New DataSet
                'Dim CLIDsTable As DataTable
                'Dim myDataColumn As DataColumn
                'CLIDsTable = New Data.DataTable("CLIDTable")

                'myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
                'CLIDsTable.Columns.Add(myDataColumn)
                'CLIDs.Tables.Add(CLIDsTable)

                AddInvMaster(MRListData, OracleLoginData)
                MRListData = CLIDforMR(MRListData, OracleLoginData)
                MRListData = PRODforMR(MRListData, OracleLoginData)

                PrintCLIDforMR(MRListData, Printer)

                Return MRListData
            Else
                Return MRListData
            End If
        Catch ex As Exception
            ErrorLogging("InvMigration-PostMR", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Private Function AddInvMaster(ByVal MRListData As DataSet, ByVal OracleLoginData As ERPLogin)
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim InvMasterSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim strCMD As String
            Dim i As Integer

            myConn.Open()
            For i = 0 To MRListData.Tables("MRData").Rows.Count - 1
                MRListData.Tables("MRData").Rows(i)("DateCode") = Mid(MRListData.Tables("MRData").Rows(i)("DateCode"), 1, 20)
                MRListData.Tables("MRData").Rows(i)("LotNo") = Mid(MRListData.Tables("MRData").Rows(i)("LotNo"), 1, 20)

                If MRListData.Tables("MRData").Rows(i)("ExpiredDate") = "" Then
                    strCMD = String.Format("INSERT INTO T_InvMaster (InvID,Line,Item,Revision,Subinventory,Locator,UnitQty,NoofPackage,Manufacturer,ManufacturerPN, CreatedOn, CreatedBy, DateCode, LotNo, BPCSPN) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',getDate(),'{10}','{11}','{12}','{13}')", MRListData.Tables("MRData").Rows(i)("InvID"), MRListData.Tables("MRData").Rows(i)("Line"), MRListData.Tables("MRData").Rows(i)("Item"), MRListData.Tables("MRData").Rows(i)("Revision"), MRListData.Tables("MRData").Rows(i)("Subinventory"), MRListData.Tables("MRData").Rows(i)("Locator"), MRListData.Tables("MRData").Rows(i)("UnitQty"), MRListData.Tables("MRData").Rows(i)("NoofPackage"), MRListData.Tables("MRData").Rows(i)("Manufacturer"), MRListData.Tables("MRData").Rows(i)("ManufacturerPN"), OracleLoginData.User.ToUpper, MRListData.Tables("MRData").Rows(i)("DateCode"), MRListData.Tables("MRData").Rows(i)("LotNo"), MRListData.Tables("MRData").Rows(i)("BPCSPN"))
                Else
                    strCMD = String.Format("INSERT INTO T_InvMaster (InvID,Line,Item,Revision,Subinventory,Locator,ExpiredDate,UnitQty,NoofPackage,Manufacturer,ManufacturerPN, CreatedOn, CreatedBy, DateCode, LotNo, BPCSPN) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',getDate(),'{11}','{12}','{13}','{14}')", MRListData.Tables("MRData").Rows(i)("InvID"), MRListData.Tables("MRData").Rows(i)("Line"), MRListData.Tables("MRData").Rows(i)("Item"), MRListData.Tables("MRData").Rows(i)("Revision"), MRListData.Tables("MRData").Rows(i)("Subinventory"), MRListData.Tables("MRData").Rows(i)("Locator"), MRListData.Tables("MRData").Rows(i)("ExpiredDate"), MRListData.Tables("MRData").Rows(i)("UnitQty"), MRListData.Tables("MRData").Rows(i)("NoofPackage"), MRListData.Tables("MRData").Rows(i)("Manufacturer"), MRListData.Tables("MRData").Rows(i)("ManufacturerPN"), OracleLoginData.User.ToUpper, MRListData.Tables("MRData").Rows(i)("DateCode"), MRListData.Tables("MRData").Rows(i)("LotNo"), MRListData.Tables("MRData").Rows(i)("BPCSPN"))
                End If
                InvMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                InvMasterSQLCommand.CommandTimeout = TimeOut_M5
                ra = InvMasterSQLCommand.ExecuteNonQuery()
            Next
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("InvMigration-AddInvMaster", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Private Function CLIDforMR(ByVal MRListData As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim i As Integer
            Dim NewCLIDCommand As SqlClient.SqlCommand
            Dim InvLabelSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim rb As Integer
            Dim strCLID As String
            Dim NextCLID As String
            Dim myDataRow As Data.DataRow

            myConn.Open()
            For i = 0 To MRListData.Tables("CLIDData").Rows.Count - 1
                If FixNull(MRListData.Tables("CLIDData").Rows(i)("RTLot")) <> "" AndAlso FixNull(MRListData.Tables("CLIDData").Rows(i)("ExpDate")) <> "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@RTLot,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ExpDate,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                    NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
                    NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
                    NewCLIDCommand.Parameters("@RTLot").Value = MRListData.Tables("CLIDData").Rows(i)("RTLot")
                    NewCLIDCommand.Parameters("@ExpDate").Value = MRListData.Tables("CLIDData").Rows(i)("ExpDate")
                ElseIf FixNull(MRListData.Tables("CLIDData").Rows(i)("RTLot")) <> "" AndAlso FixNull(MRListData.Tables("CLIDData").Rows(i)("ExpDate")) = "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@RTLot,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                    NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
                    NewCLIDCommand.Parameters("@RTLot").Value = MRListData.Tables("CLIDData").Rows(i)("RTLot")
                ElseIf FixNull(MRListData.Tables("CLIDData").Rows(i)("RTLot")) = "" AndAlso FixNull(MRListData.Tables("CLIDData").Rows(i)("ExpDate")) <> "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ExpDate,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                    NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
                    NewCLIDCommand.Parameters("@ExpDate").Value = MRListData.Tables("CLIDData").Rows(i)("ExpDate")
                ElseIf FixNull(MRListData.Tables("CLIDData").Rows(i)("RTLot")) = "" AndAlso FixNull(MRListData.Tables("CLIDData").Rows(i)("ExpDate")) = "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                End If
                NewCLIDCommand.Parameters.Add("@NewCLID", SqlDbType.VarChar, 20, "CLID")
                NewCLIDCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20, "OrgCode")
                NewCLIDCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 50, "MaterialNo")
                NewCLIDCommand.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 20, "MaterialRevision")
                NewCLIDCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
                NewCLIDCommand.Parameters.Add("@UOM", SqlDbType.VarChar, 10, "UOM")
                NewCLIDCommand.Parameters.Add("@QtyBaseUOM", SqlDbType.Decimal, 13, "QtyBaseUOM")
                NewCLIDCommand.Parameters.Add("@BaseUOM", SqlDbType.VarChar, 10, "BaseUOM")
                NewCLIDCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 20, "DateCode")
                NewCLIDCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 20, "LotNo")
                NewCLIDCommand.Parameters.Add("@COO", SqlDbType.VarChar, 20, "COO")
                NewCLIDCommand.Parameters.Add("@RecDocNo", SqlDbType.VarChar, 50, "RecDocNo")
                NewCLIDCommand.Parameters.Add("@RecDocItem", SqlDbType.VarChar, 10, "RecDocItem")
                'NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
                NewCLIDCommand.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
                NewCLIDCommand.Parameters.Add("@Printed", SqlDbType.VarChar, 100, "Printed")
                NewCLIDCommand.Parameters.Add("@VendorID", SqlDbType.VarChar, 10, "VendorID")
                NewCLIDCommand.Parameters.Add("@RecDate", SqlDbType.SmallDateTime, 4, "RecDate")
                NewCLIDCommand.Parameters.Add("@RoHS", SqlDbType.VarChar, 10, "RoHS")
                NewCLIDCommand.Parameters.Add("@PurOrdNo", SqlDbType.VarChar, 20, "PurOrdNo")
                NewCLIDCommand.Parameters.Add("@PurOrdItem", SqlDbType.VarChar, 10, "PurOrdItem")
                NewCLIDCommand.Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 25, "InvoiceNo")
                NewCLIDCommand.Parameters.Add("@BillofLading", SqlDbType.VarChar, 25, "BillofLading")
                NewCLIDCommand.Parameters.Add("@DN", SqlDbType.VarChar, 25, "DN")
                NewCLIDCommand.Parameters.Add("@HeaderText", SqlDbType.VarChar, 25, "HeaderText")
                'NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
                NewCLIDCommand.Parameters.Add("@ProdDate", SqlDbType.SmallDateTime, 4, "ProdDate")
                NewCLIDCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
                NewCLIDCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 10, "StockType")
                NewCLIDCommand.Parameters.Add("@MaterialDesc", SqlDbType.VarChar, 50, "MaterialDesc")
                NewCLIDCommand.Parameters.Add("@VendorName", SqlDbType.VarChar, 50, "VendorName")
                NewCLIDCommand.Parameters.Add("@VendorPN", SqlDbType.VarChar, 50, "VendorPN")
                NewCLIDCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
                NewCLIDCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
                NewCLIDCommand.Parameters.Add("@Operator", SqlDbType.VarChar, 50, "Operator")
                NewCLIDCommand.Parameters.Add("@IsTraceable", SqlDbType.VarChar, 10, "IsTraceable")
                NewCLIDCommand.Parameters.Add("@MatDocNo", SqlDbType.VarChar, 50, "MatDocNo")
                NewCLIDCommand.Parameters.Add("@MatDocItem", SqlDbType.VarChar, 10, "MatDocItem")
                NewCLIDCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 50, "Manufacturer")
                NewCLIDCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 50, "ManufacturerPN")
                NewCLIDCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
                NewCLIDCommand.Parameters.Add("@AddlData", SqlDbType.VarChar, 20, "AddlData")
                NewCLIDCommand.Parameters.Add("@Stemp", SqlDbType.VarChar, 50, "Stemp")
                NewCLIDCommand.Parameters.Add("@MSL", SqlDbType.VarChar, 50, "MSL")
                NewCLIDCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 50, "LastTransaction")

                NextCLID = GetNextCLID(OracleLoginData)
                NewCLIDCommand.Parameters("@NewCLID").Value = NextCLID
                MRListData.Tables("CLIDData").Rows(i)("NewCLID") = NextCLID
                NewCLIDCommand.Parameters("@OrgCode").Value = MRListData.Tables("CLIDData").Rows(i)("OrgCode")
                NewCLIDCommand.Parameters("@MaterialNo").Value = MRListData.Tables("CLIDData").Rows(i)("MaterialNo")
                NewCLIDCommand.Parameters("@MaterialRevision").Value = MRListData.Tables("CLIDData").Rows(i)("MaterialRevision")
                NewCLIDCommand.Parameters("@Qty").Value = MRListData.Tables("CLIDData").Rows(i)("Qty")
                NewCLIDCommand.Parameters("@UOM").Value = MRListData.Tables("CLIDData").Rows(i)("UOM")
                NewCLIDCommand.Parameters("@QtyBaseUOM").Value = MRListData.Tables("CLIDData").Rows(i)("QtyBaseUOM")
                NewCLIDCommand.Parameters("@BaseUOM").Value = MRListData.Tables("CLIDData").Rows(i)("BaseUOM")
                NewCLIDCommand.Parameters("@DateCode").Value = MRListData.Tables("CLIDData").Rows(i)("DateCode")
                NewCLIDCommand.Parameters("@LotNo").Value = MRListData.Tables("CLIDData").Rows(i)("LotNo")
                NewCLIDCommand.Parameters("@COO").Value = MRListData.Tables("CLIDData").Rows(i)("COO")
                NewCLIDCommand.Parameters("@RecDocNo").Value = MRListData.Tables("CLIDData").Rows(i)("RecDocNo")
                NewCLIDCommand.Parameters("@RecDocItem").Value = MRListData.Tables("CLIDData").Rows(i)("RecDocItem")
                'NewCLIDCommand.Parameters("@RTLot").Value = MRListData.Tables("CLIDData").Rows(i)("RTLot")
                NewCLIDCommand.Parameters("@CreatedBy").Value = MRListData.Tables("CLIDData").Rows(i)("CreatedBy")
                NewCLIDCommand.Parameters("@Printed").Value = MRListData.Tables("CLIDData").Rows(i)("Printed")
                NewCLIDCommand.Parameters("@VendorID").Value = MRListData.Tables("CLIDData").Rows(i)("VendorID")
                NewCLIDCommand.Parameters("@RecDate").Value = MRListData.Tables("CLIDData").Rows(i)("RecDate")
                NewCLIDCommand.Parameters("@RoHS").Value = MRListData.Tables("CLIDData").Rows(i)("RoHS")
                NewCLIDCommand.Parameters("@PurOrdNo").Value = MRListData.Tables("CLIDData").Rows(i)("PurOrdNo")
                NewCLIDCommand.Parameters("@PurOrdItem").Value = MRListData.Tables("CLIDData").Rows(i)("PurOrdItem")
                NewCLIDCommand.Parameters("@InvoiceNo").Value = MRListData.Tables("CLIDData").Rows(i)("InvoiceNo")
                NewCLIDCommand.Parameters("@BillofLading").Value = MRListData.Tables("CLIDData").Rows(i)("BillofLading")
                NewCLIDCommand.Parameters("@DN").Value = MRListData.Tables("CLIDData").Rows(i)("DN")
                NewCLIDCommand.Parameters("@HeaderText").Value = MRListData.Tables("CLIDData").Rows(i)("HeaderText")
                'NewCLIDCommand.Parameters("@ExpDate").Value = MRListData.Tables("CLIDData").Rows(i)("ExpDate")
                NewCLIDCommand.Parameters("@ProdDate").Value = MRListData.Tables("CLIDData").Rows(i)("ProdDate")
                NewCLIDCommand.Parameters("@ReasonCode").Value = MRListData.Tables("CLIDData").Rows(i)("ReasonCode")
                NewCLIDCommand.Parameters("@StockType").Value = MRListData.Tables("CLIDData").Rows(i)("StockType")
                NewCLIDCommand.Parameters("@MaterialDesc").Value = SQLString(MRListData.Tables("CLIDData").Rows(i)("MaterialDesc"))
                NewCLIDCommand.Parameters("@VendorName").Value = MRListData.Tables("CLIDData").Rows(i)("VendorName")
                NewCLIDCommand.Parameters("@VendorPN").Value = MRListData.Tables("CLIDData").Rows(i)("VendorPN")
                NewCLIDCommand.Parameters("@SLOC").Value = MRListData.Tables("CLIDData").Rows(i)("SLOC")
                NewCLIDCommand.Parameters("@StorageBin").Value = MRListData.Tables("CLIDData").Rows(i)("StorageBin")
                NewCLIDCommand.Parameters("@Operator").Value = MRListData.Tables("CLIDData").Rows(i)("Operator")
                NewCLIDCommand.Parameters("@IsTraceable").Value = MRListData.Tables("CLIDData").Rows(i)("IsTraceable")
                NewCLIDCommand.Parameters("@MatDocNo").Value = MRListData.Tables("CLIDData").Rows(i)("MatDocNo")
                NewCLIDCommand.Parameters("@MatDocItem").Value = MRListData.Tables("CLIDData").Rows(i)("MatDocItem")
                NewCLIDCommand.Parameters("@Manufacturer").Value = MRListData.Tables("CLIDData").Rows(i)("Manufacturer")
                NewCLIDCommand.Parameters("@ManufacturerPN").Value = MRListData.Tables("CLIDData").Rows(i)("ManufacturerPN")
                NewCLIDCommand.Parameters("@QMLStatus").Value = MRListData.Tables("CLIDData").Rows(i)("QMLStatus")
                NewCLIDCommand.Parameters("@AddlData").Value = MRListData.Tables("CLIDData").Rows(i)("AddlData")
                NewCLIDCommand.Parameters("@Stemp").Value = MRListData.Tables("CLIDData").Rows(i)("Stemp")
                NewCLIDCommand.Parameters("@MSL").Value = MRListData.Tables("CLIDData").Rows(i)("MSL")
                NewCLIDCommand.Parameters("@LastTransaction").Value = "InventoryMigration"

                NewCLIDCommand.CommandTimeout = TimeOut_M5
                NewCLIDCommand.CommandType = CommandType.Text
                ra = NewCLIDCommand.ExecuteNonQuery()

                myDataRow = MRListData.Tables("CLIDTable").NewRow()
                myDataRow("InvID") = MRListData.Tables("CLIDData").Rows(i)("InvID")
                myDataRow("Line") = MRListData.Tables("CLIDData").Rows(i)("Line")
                myDataRow("CLID") = NextCLID
                MRListData.Tables("CLIDTable").Rows.Add(myDataRow)

                strCLID = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", MRListData.Tables("CLIDData").Rows(i)("InvID"), MRListData.Tables("CLIDData").Rows(i)("Line"), NextCLID)
                InvLabelSQLCommand = New SqlClient.SqlCommand(strCLID, myConn)
                InvLabelSQLCommand.CommandTimeout = TimeOut_M5
                rb = InvLabelSQLCommand.ExecuteNonQuery()

            Next
            myConn.Close()
            Return MRListData
        Catch ex As Exception
            ErrorLogging("InvMigration-AddCLIDforMR", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Private Function PRODforMR(ByVal MRListData As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim myDataRow As Data.DataRow
            Dim CLMasterSQLCommand As SqlClient.SqlCommand
            Dim InvLabelSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim rb As Integer
            Dim strCMD As String
            Dim strProd As String
            Dim ProdCLID, OrgCode, matlNo, matlDesc, UOM, PurOrdNo, statusCode, RTLot, ExpDate, DateCode, LotNo, Revision, CreatedBy, SubInv, Location As String
            Dim QtyBaseUOM As Decimal
            Dim i As Integer
            Dim result1 As New Object
            Dim Last_Transaction As String

            myConn.Open()
            For i = 0 To MRListData.Tables("ProdData").Rows.Count - 1
                ProdCLID = GetNextProdID(OracleLoginData)
                MRListData.Tables("ProdData").Rows(i)("CLID") = ProdCLID
                OrgCode = MRListData.Tables("ProdData").Rows(i)("OrgCode")
                matlNo = MRListData.Tables("ProdData").Rows(i)("MaterialNo")
                matlDesc = MRListData.Tables("ProdData").Rows(i)("MaterialDesc")
                UOM = MRListData.Tables("ProdData").Rows(i)("BaseUOM")
                PurOrdNo = MRListData.Tables("ProdData").Rows(i)("PurOrdNo")
                QtyBaseUOM = MRListData.Tables("ProdData").Rows(i)("QtyBaseUOM")
                'If Not MRListData.Tables("ProdData").Rows(i)("RTLot") Is DBNull.Value Then
                '    RTLot = MRListData.Tables("ProdData").Rows(i)("RTLot")
                'End If
                'If Not MRListData.Tables("ProdData").Rows(i)("ExpDate") Is DBNull.Value Then
                '    ExpDate = MRListData.Tables("ProdData").Rows(i)("ExpDate")
                'End If
                DateCode = MRListData.Tables("ProdData").Rows(i)("DateCode")
                LotNo = MRListData.Tables("ProdData").Rows(i)("LotNo")
                statusCode = "1"
                CreatedBy = OracleLoginData.User.ToUpper
                Revision = MRListData.Tables("ProdData").Rows(i)("MaterialRevision")
                SubInv = MRListData.Tables("ProdData").Rows(i)("SubInv")
                Location = MRListData.Tables("ProdData").Rows(i)("Locator")
                Last_Transaction = "InventoryMigration"

                If FixNull(MRListData.Tables("ProdData").Rows(i)("RTLot")) = "" AndAlso FixNull(MRListData.Tables("ProdData").Rows(i)("ExpDate")) = "" Then
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,PurOrdNo,QtyBaseUOM,DateCode,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,LastTransaction,MaterialDesc) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}','{10}','{11}',getDate(),getDate(),'{12}','{13}','{14}','{15}','{16}')", ProdCLID, OrgCode, matlNo, Revision, QtyBaseUOM, UOM, PurOrdNo, QtyBaseUOM, DateCode, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, Last_Transaction, matlDesc)
                    'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,LastTransaction) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}')", ProdCLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", MRListData.Tables("ProdData").Rows(i)("InvID"), MRListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                ElseIf FixNull(MRListData.Tables("ProdData").Rows(i)("RTLot")) = "" AndAlso FixNull(MRListData.Tables("ProdData").Rows(i)("ExpDate")) <> "" Then
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,PurOrdNo,QtyBaseUOM,DateCode,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,ExpDate,CreatedBy,SLOC,StorageBin,LastTransaction,MaterialDesc) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}','{10}','{11}',getDate(),getDate(),'{12}','{13}','{14}',,'{15}','{16}','{17}')", ProdCLID, OrgCode, matlNo, Revision, QtyBaseUOM, UOM, PurOrdNo, QtyBaseUOM, DateCode, LotNo, statusCode, UOM, MRListData.Tables("ProdData").Rows(i)("ExpDate"), CreatedBy, SubInv, Location, Last_Transaction, matlDesc)
                    'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,LastTransaction) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}')", ProdCLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", MRListData.Tables("ProdData").Rows(i)("InvID"), MRListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                ElseIf FixNull(MRListData.Tables("ProdData").Rows(i)("RTLot")) <> "" AndAlso FixNull(MRListData.Tables("ProdData").Rows(i)("ExpDate")) = "" Then
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,PurOrdNo,QtyBaseUOM,DateCode,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,RTLot,CreatedBy,SLOC,StorageBin,LastTransaction,MaterialDesc) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}','{10}','{11}',getDate(),getDate(),'{12}','{13}','{14}','{15}','{16}','{17}')", ProdCLID, OrgCode, matlNo, Revision, QtyBaseUOM, UOM, PurOrdNo, QtyBaseUOM, DateCode, LotNo, statusCode, UOM, MRListData.Tables("ProdData").Rows(i)("RTLot"), CreatedBy, SubInv, Location, Last_Transaction, matlDesc)
                    'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,LastTransaction) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}')", ProdCLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", MRListData.Tables("ProdData").Rows(i)("InvID"), MRListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                ElseIf FixNull(MRListData.Tables("ProdData").Rows(i)("RTLot")) <> "" AndAlso FixNull(MRListData.Tables("ProdData").Rows(i)("ExpDate")) <> "" Then
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,PurOrdNo,QtyBaseUOM,DateCode,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,ExpDate,RTLot,CreatedBy,SLOC,StorageBin,LastTransaction,MaterialDesc) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}','{10}','{11}',getDate(),getDate(),'{12}','{13}','{14}','{15}','{16}','{17}','{18}')", ProdCLID, OrgCode, matlNo, Revision, QtyBaseUOM, UOM, PurOrdNo, QtyBaseUOM, DateCode, LotNo, statusCode, UOM, MRListData.Tables("ProdData").Rows(i)("ExpDate"), MRListData.Tables("ProdData").Rows(i)("RTLot"), CreatedBy, SubInv, Location, Last_Transaction, matlDesc)
                    'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,LastTransaction) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}')", ProdCLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", MRListData.Tables("ProdData").Rows(i)("InvID"), MRListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                End If

                'myConn.Open()
                CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                InvLabelSQLCommand = New SqlClient.SqlCommand(strProd, myConn)
                InvLabelSQLCommand.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand.ExecuteNonQuery()
                rb = InvLabelSQLCommand.ExecuteNonQuery()
                'myConn.Close()

                myDataRow = MRListData.Tables("CLIDTable").NewRow()
                myDataRow("InvID") = MRListData.Tables("ProdData").Rows(i)("InvID")
                myDataRow("Line") = MRListData.Tables("ProdData").Rows(i)("Line")
                myDataRow("CLID") = ProdCLID
                MRListData.Tables("CLIDTable").Rows.Add(myDataRow)
            Next
            myConn.Close()

            Return MRListData
        Catch ex As Exception
            ErrorLogging("InvMigration-AddProdIDforMR", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function PrintCLIDforMR(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        PrintCLIDforMR = True
        Try
            If CLIDs Is Nothing OrElse CLIDs.Tables.Count = 0 OrElse CLIDs.Tables("CLIDTable").Rows.Count = 0 Then
                PrintCLIDforMR = False
                Exit Function
            End If

            'Sort CLID by Ascending before written to eTrace table T_LabelPrint
            Dim dtPrint As DataTable = New DataTable
            Dim SortColName As String = CLIDs.Tables("CLIDTable").Columns(2).ColumnName
            SortColName = SortColName & " ASC"

            CLIDs.Tables("CLIDTable").DefaultView.Sort = SortColName
            dtPrint = CLIDs.Tables("CLIDTable").DefaultView.ToTable()

            CLIDs = New DataSet
            CLIDs.Tables.Add(dtPrint)

            Dim i As Integer
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                If PrintCLID(CLIDs.Tables(0).Rows(i)("CLID"), Printer) = False Then
                    PrintCLIDforMR = False
                End If
                Sleep(5)
            Next
        Catch ex As Exception
            ErrorLogging("InvMigration-PrintCLIDforMR", "", ex.Message & ex.Source, "E")
            PrintCLIDforMR = False
        End Try
    End Function

    Public Function CS_GetCLIDInfo(ByVal clid As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_CS_GetCLIDInfo '{0}', '{1}'", clid, OracleLoginData.UserID)
                ds = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("HH-CS_GetCLIDInfo", "", "CLID: " & clid & ", " & ex.ToString, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function CS_UpdateCLID(ByVal clid As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sqlstr As String
            Try
                sqlstr = String.Format("Update T_CLMaster SET StatusCode = '1' where CLID = '{0}'", clid)
                da.ExecuteNonQuery(sqlstr)
                CS_UpdateCLID = ""
            Catch ex As Exception
                CS_UpdateCLID = ex.Message.ToString
                ErrorLogging("HH-CS_GetCLIDInfo", "", "CLID: " & clid & ", " & ex.ToString, "E")
            End Try
        End Using
    End Function

    Public Function GetBerth() As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_getBerth")
                DS = da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("GetBerth", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function

    Public Function LEDDashBoardByRack() As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_LEDDashBoardByRack")
                DS = da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("LEDDashBoardByRack", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function
    Public Function LEDDashBoardPCB(ByVal PCBWarehouse As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_LEDDashBoardPCB '{0}'", PCBWarehouse)
                DS = da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("LEDDashBoardPCB", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function

    Public Function GetDashboardData() As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_GetDashboardData")
                DS = da.ExecuteDataSet(strSQL, 1)
            Catch ex As Exception
                ErrorLogging("GetDashboardData", "", ex.Message & ex.Source, "E")
            End Try
            Return DS
        End Using
    End Function

    Public Function ValidBatchData(ByVal BatchListData As DataSet, ByVal dsAML As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Dim MatList, SapSLOC As DataSet
            Dim i, j, m, n As Integer
            Dim Flag_Rev, Flag_Sub As Boolean
            Dim MaterialNo, Manufacturer, ManufacturerPN As String
            Dim Rev, RevList(), RstSubList() As String
            Dim DR() As DataRow = Nothing
            Dim DR_SLOC(), DR_SLOCType(), DR_SLOCTypeBin(), DR_AML(), DR_MFR_MPN(), DR_Item() As DataRow
            Dim PostDataRow, CLIDDataRow, ProdDataRow As Data.DataRow
            Dim s_date, ExpDate, itemtype, LotControl, LifeControl, RTLot As String
            Dim nolife_flag As Boolean
            Dim LifeDays As Integer

            MatList = New DataSet
            SapSLOC = New DataSet
            Using da As DataAccess = GetDataAccess()
                Try
                    Dim Sqlstr As String

                    Sqlstr = String.Format("SELECT * from T_SAPPN with (nolock) where not ConvFactor is NULL ")
                    MatList = da.ExecuteDataSet(Sqlstr, "MatData")

                    Sqlstr = String.Format("SELECT * from T_SAPSLOC with (nolock) where not Subinv is NULL ")
                    SapSLOC = da.ExecuteDataSet(Sqlstr, "SLOC")

                Catch ex As Exception
                    ErrorLogging("ZSMigration-ValidBatchData-GetSourceData", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                    Return Nothing
                End Try
            End Using

            For i = 0 To BatchListData.Tables("BatchData").Rows.Count - 1
                DR = Nothing
                DR_SLOC = Nothing
                DR_SLOCType = Nothing
                DR_SLOCTypeBin = Nothing
                DR_AML = Nothing
                DR_MFR_MPN = Nothing
                DR_AML = Nothing
                DR_Item = Nothing
                itemtype = ""
                nolife_flag = False

                Flag_Rev = False
                Flag_Sub = False
                MaterialNo = BatchListData.Tables("BatchData").Rows(i)("Item")
                DR = MatList.Tables(0).Select(" SAPPN = '" & MaterialNo & "'")
                If Not DR.Length > 0 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Error: Material is invalid"
                    Continue For
                End If
                If DR(0)("RevControl") = True Then
                    Erase RevList
                    If InStr(DR(0)("RevList"), ",") > 0 Then
                        RevList = Split(DR(0)("RevList"), ",")
                        Rev = RevList(0)
                        For j = LBound(RevList) To UBound(RevList)
                            If Rev.CompareTo(RevList(j)) = 0 Or Rev.CompareTo(RevList(j)) = 1 Then
                            ElseIf Rev.CompareTo(RevList(j)) = -1 Then
                                Rev = RevList(j)
                            End If
                        Next
                    Else
                        Rev = DR(0)("RevList")
                    End If
                ElseIf DR(0)("RevControl") = False Then
                    Rev = ""
                End If

                DR_SLOC = SapSLOC.Tables(0).Select(" SLOC = '" & BatchListData.Tables("BatchData").Rows(i)("SLOC") & "'")
                If Not DR_SLOC.Length > 0 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Error: SLOC is invalid"
                    Continue For
                End If

                DR_SLOCType = SapSLOC.Tables(0).Select(" SLOC = '" & BatchListData.Tables("BatchData").Rows(i)("SLOC") & "'and SType = '" & FixNull(BatchListData.Tables("BatchData").Rows(i)("Subinventory")) & "'")

                If Not DR_SLOCType.Length > 0 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Error: SType is invalid"
                    Continue For
                End If

                DR_SLOCTypeBin = SapSLOC.Tables(0).Select(" SLOC = '" & BatchListData.Tables("BatchData").Rows(i)("SLOC") & "'and SType = '" & FixNull(BatchListData.Tables("BatchData").Rows(i)("Subinventory")) & "'and Bin = '" & FixNull(BatchListData.Tables("BatchData").Rows(i)("Locator")) & "'")

                If Not DR_SLOCTypeBin.Length > 0 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Error: SBin is invalid"
                    Continue For
                End If

                If DR(0)("ExpControl") = True AndAlso FixNull(BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")) = "" Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Error: ExpiredDate is required"
                    Continue For
                End If
                If DR(0)("ExpControl") = False AndAlso FixNull(BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")) <> "" Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = "Warning: No shelf life control, entered ExpiredDate won't be considered."
                End If

                If DR(0)("ExpControl") = True Then
                    ExpDate = Format(CDate(BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")), "yyyyMMdd")
                Else
                    ExpDate = ""
                End If
                RTLot = ExpDate & BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")

                Manufacturer = FixNull(BatchListData.Tables("BatchData").Rows(i)("Manufacturer"))
                If Manufacturer <> "" Then
                    Manufacturer = SQLString(Manufacturer)
                End If
                ManufacturerPN = FixNull(BatchListData.Tables("BatchData").Rows(i)("ManufacturerPN"))
                If ManufacturerPN <> "" Then
                    ManufacturerPN = SQLString(ManufacturerPN)
                End If


                DR_MFR_MPN = dsAML.Tables("AMLData").Select(" MaterialNo = '" & MaterialNo & "' and MFR = '" & Manufacturer & "' and MPN = '" & ManufacturerPN & "'")
                If Not DR_MFR_MPN.Length > 0 AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") & "Warning: MFR & MPN not match iProd information."
                End If

                If Len(Trim(BatchListData.Tables("BatchData").Rows(i)("DateCode"))) > 20 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") & "Warning: Length of DateCode shouldn't exceed 20."
                End If
                If Len(Trim(BatchListData.Tables("BatchData").Rows(i)("LotNo"))) > 20 Then
                    BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") = BatchListData.Tables("BatchData").Rows(i)("ErrorMsg") & "Warning: Length of LotNo shouldn't exceed 20."
                End If

                If DR(0)("ItemType") = "RM" Then
                    For m = 0 To BatchListData.Tables("BatchData").Rows(i)("NoofPackage") - 1
                        'Add row for CLIDData
                        'Dim Exp_Date As String
                        Dim RecDate As String
                        ' Get Traceability Level and SafeCode according to Material Group              'Add by Yudy 11/28/2008
                        'Dim SafeCode As String
                        'Dim TraceLevel As MGTraceData
                        '            LabelForMiscReceipt.TraceLevel = "PT"    ' "NT" not needed   "CT" must needed  "PT" optional
                        'SafeCode = ""

                        RecDate = Date.Now

                        'Exp_Date = BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")
                        'If Exp_Date Is Nothing Then
                        '    Exp_Date = Replace(DateTime.Now.AddYears(3), "-", "/")
                        'End If

                        DR_AML = dsAML.Tables("AMLData").Select(" MaterialNo = '" & MaterialNo & "' ")
                        DR_Item = dsAML.Tables("ItemData").Select(" MaterialNo = '" & MaterialNo & "' ")

                        CLIDDataRow = BatchListData.Tables("CLIDData").NewRow()
                        CLIDDataRow("InvID") = BatchListData.Tables("BatchData").Rows(i)("InvID")
                        CLIDDataRow("Line") = BatchListData.Tables("BatchData").Rows(i)("Line")
                        CLIDDataRow("NewCLID") = ""
                        CLIDDataRow("OrgCode") = OracleLoginData.OrgCode
                        CLIDDataRow("MaterialNo") = MaterialNo
                        CLIDDataRow("MaterialRevision") = Rev
                        CLIDDataRow("Qty") = BatchListData.Tables("BatchData").Rows(i)("UnitQty")
                        CLIDDataRow("UOM") = DR(0)("SAPUOM")
                        CLIDDataRow("QtyBaseUOM") = BatchListData.Tables("BatchData").Rows(i)("UnitQty")
                        CLIDDataRow("BaseUOM") = DR(0)("SAPUOM")
                        CLIDDataRow("DateCode") = Mid(BatchListData.Tables("BatchData").Rows(i)("DateCode"), 1, 20)
                        CLIDDataRow("LotNo") = Mid(BatchListData.Tables("BatchData").Rows(i)("LotNo"), 1, 20)
                        CLIDDataRow("COO") = DBNull.Value
                        CLIDDataRow("RecDocYear") = Date.Now.Year
                        CLIDDataRow("RecDocNo") = BatchListData.Tables("BatchData").Rows(i)("InvID")   'BatchListData.Tables("BatchData").Rows(i)("InvID")
                        CLIDDataRow("RecDocItem") = DBNull.Value
                        CLIDDataRow("CreatedBy") = OracleLoginData.User.ToUpper
                        CLIDDataRow("Printed") = "True"
                        CLIDDataRow("VendorID") = ""
                        CLIDDataRow("RecDate") = CDate(RecDate)
                        If DR_Item.Length > 0 Then
                            CLIDDataRow("RoHS") = DR_Item(0)("RoHS")
                        End If
                        CLIDDataRow("PurOrdNo") = ""
                        CLIDDataRow("PurOrdItem") = ""
                        CLIDDataRow("InvoiceNo") = ""
                        CLIDDataRow("BillofLading") = ""
                        CLIDDataRow("DN") = ""
                        CLIDDataRow("HeaderText") = ""
                        'If FixNull(Exp_Date) = "" Then
                        '    CLIDDataRow("ExpDate") = DateTime.Now.AddYears(3)   'DBNull.Value
                        'Else
                        '    CLIDDataRow("ExpDate") = CDate(Exp_Date)
                        'End If
                        'CLIDDataRow("ExpDate") = CDate(Exp_Date)
                        'If FixNull(DR(0)("o_lot_control_code")) = 2 Then
                        '    CLIDDataRow("RTLot") = ExpDate & BatchListData.Tables("BatchData").Rows(i)("InvID")  'BatchListData.Tables("BatchData").Rows(i)("InvID")
                        '    CLIDDataRow("ExpDate") = PostDataRow("p_lot_expiration_date")

                        '    CLIDDataRow("RTLot") = DBNull.Value
                        '    CLIDDataRow("ExpDate") = DBNull.Value
                        'End If

                        If DR(0)("ExpControl") = True AndAlso FixNull(BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")) <> "" Then
                            CLIDDataRow("ExpDate") = BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")
                        Else
                            CLIDDataRow("ExpDate") = DBNull.Value
                        End If
                        CLIDDataRow("ProdDate") = DBNull.Value
                        CLIDDataRow("ReasonCode") = ""
                        CLIDDataRow("StockType") = "QP"
                        CLIDDataRow("MaterialDesc") = SQLString(DR(0)("Description"))
                        CLIDDataRow("VendorName") = ""
                        CLIDDataRow("VendorPN") = ""
                        CLIDDataRow("SLOC") = BatchListData.Tables("BatchData").Rows(i)("SLOC")
                        CLIDDataRow("StorageType") = BatchListData.Tables("BatchData").Rows(i)("Subinventory")
                        CLIDDataRow("StorageBin") = BatchListData.Tables("BatchData").Rows(i)("Locator")
                        CLIDDataRow("Operator") = OracleLoginData.User.ToUpper
                        CLIDDataRow("IsTraceable") = "PT"
                        CLIDDataRow("MatDocNo") = BatchListData.Tables("BatchData").Rows(i)("InvID")
                        CLIDDataRow("MatDocItem") = DBNull.Value
                        If InStr(Manufacturer, "^") > 0 Then
                            CLIDDataRow("Manufacturer") = Replace(Manufacturer, "^", "~")
                        Else
                            CLIDDataRow("Manufacturer") = Manufacturer
                        End If
                        If InStr(ManufacturerPN, "^") > 0 Then
                            CLIDDataRow("ManufacturerPN") = Replace(ManufacturerPN, "^", "~")
                        Else
                            CLIDDataRow("ManufacturerPN") = ManufacturerPN
                        End If
                        'CLIDDataRow("Manufacturer") = Manufacturer
                        'CLIDDataRow("ManufacturerPN") = ManufacturerPN
                        If DR_AML.Length > 0 Then
                            CLIDDataRow("QMLStatus") = DR_AML(0)("AMLStatus")
                        End If
                        If DR_Item.Length > 0 Then
                            CLIDDataRow("AddlData") = DR_Item(0)("AddlData")
                            CLIDDataRow("Stemp") = DR_Item(0)("Stemp")
                            CLIDDataRow("MSL") = DR_Item(0)("MSL")
                            'CLIDDataRow("MSL") = DR_Item(0)("MSL") & " (" & BatchListData.Tables("BatchData").Rows(i)("BPCSPN") & ")"
                        Else
                            CLIDDataRow("AddlData") = ""
                            CLIDDataRow("Stemp") = ""
                            CLIDDataRow("MSL") = ""
                            'CLIDDataRow("MSL") = DR_Item(0)("MSL") & " (" & BatchListData.Tables("BatchData").Rows(i)("BPCSPN") & ")"
                        End If
                        BatchListData.Tables("CLIDData").Rows.Add(CLIDDataRow)
                    Next
                ElseIf DR(0)("ItemType") = "FG" Or DR(0)("ItemType") = "SA" Or DR(0)("ItemType") = "PH" Then
                    For n = 0 To BatchListData.Tables("BatchData").Rows(i)("NoofPackage") - 1
                        'Add row for ProdData
                        ProdDataRow = BatchListData.Tables("ProdData").NewRow()
                        ProdDataRow("InvID") = BatchListData.Tables("BatchData").Rows(i)("InvID")
                        ProdDataRow("Line") = BatchListData.Tables("BatchData").Rows(i)("Line")
                        ProdDataRow("CLID") = ""
                        ProdDataRow("PurOrdNo") = BatchListData.Tables("BatchData").Rows(i)("InvID")
                        ProdDataRow("OrgCode") = OracleLoginData.OrgCode
                        ProdDataRow("MaterialNo") = MaterialNo
                        ProdDataRow("MaterialDesc") = SQLString(DR(0)("Description"))
                        ProdDataRow("MaterialRevision") = Rev
                        ProdDataRow("QtyBaseUOM") = BatchListData.Tables("BatchData").Rows(i)("UnitQty")
                        ProdDataRow("BaseUOM") = DR(0)("SAPUOM")

                        If DR(0)("ExpControl") = True AndAlso FixNull(BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")) <> "" Then
                            ProdDataRow("ExpDate") = BatchListData.Tables("BatchData").Rows(i)("ExpiredDate")
                        Else
                            ProdDataRow("ExpDate") = DBNull.Value
                        End If
                        ProdDataRow("RecDocYear") = Date.Now.Year
                        ProdDataRow("RecDocNo") = BatchListData.Tables("BatchData").Rows(i)("InvID")   'BatchListData.Tables("BatchData").Rows(i)("InvID")
                        ProdDataRow("DateCode") = Mid(BatchListData.Tables("BatchData").Rows(i)("DateCode"), 1, 20)
                        ProdDataRow("LotNo") = Mid(BatchListData.Tables("BatchData").Rows(i)("LotNo"), 1, 20)
                        If InStr(Manufacturer, "^") > 0 Then
                            ProdDataRow("Manufacturer") = Replace(Manufacturer, "^", "~")
                        Else
                            ProdDataRow("Manufacturer") = Manufacturer
                        End If
                        If InStr(ManufacturerPN, "^") > 0 Then
                            ProdDataRow("ManufacturerPN") = Replace(ManufacturerPN, "^", "~")
                        Else
                            ProdDataRow("ManufacturerPN") = ManufacturerPN
                        End If
                        ProdDataRow("SLOC") = BatchListData.Tables("BatchData").Rows(i)("SLOC")
                        ProdDataRow("StorageType") = BatchListData.Tables("BatchData").Rows(i)("Subinventory")
                        ProdDataRow("StorageBin") = BatchListData.Tables("BatchData").Rows(i)("Locator")

                        BatchListData.Tables("ProdData").Rows.Add(ProdDataRow)
                    Next
                End If
            Next
            Return BatchListData

        Catch ex As Exception
            ErrorLogging("InvMigration-ValidBatchData", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
            Return Nothing
        End Try

    End Function

    Public Function PostBatchList(ByVal BatchListData As DataSet, ByVal MoveType As String, ByVal Print As Boolean, ByVal Printer As String, ByVal OracleLoginData As ERPLogin) As PostBatchRslt
        Try
            Dim TransactionID As Double
            Dim ID As String
            Dim MiscType As String
            Dim BatchData_Return As New DataSet
            Dim DR() As DataRow = Nothing
            Dim CLIDs As DataSet
            'Dim eTraceGoodsMvt As eTraceGoodsMvt.eTraceGoodsMvtService = New eTraceGoodsMvt.eTraceGoodsMvtService
            'TransactionID = GetNextInvID(OracleLoginData)
            MiscType = ""
            BatchData_Return = BatchListData.Clone

            AddInvMaster_ForBatch(BatchListData, OracleLoginData)
            BatchListData = CLIDforBatch(BatchListData, OracleLoginData)
            BatchListData = PRODforBatch(BatchListData, OracleLoginData)

            CLIDs = New DataSet
            CLIDs = BatchListData.Copy
            CLIDs.Tables.Remove("BatchData")
            CLIDs.Tables.Remove("CLIDData")
            CLIDs.Tables.Remove("ProdData")
            CLIDs.Tables(0).Columns.Remove("InvID")
            CLIDs.Tables(0).Columns.Remove("Line")

            PostBatchList.PrintFlag = False
            PostBatchList.BatchList = New DataSet

            PostBatchList.BatchList = BatchListData
            'If Print = True Then
            '    eTraceGoodsMvt.Timeout = 1000 * 60 * 120
            '    PostBatchList.PrintFlag = eTraceGoodsMvt.PrintCLIDs(CLIDs, Printer, True)
            'End If

            Return PostBatchList

        Catch ex As Exception
            ErrorLogging("ZSMigration-PostBatchList", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Private Function AddInvMaster_ForBatch(ByVal BatchListData As DataSet, ByVal OracleLoginData As ERPLogin)
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim InvMasterSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim strCMD As String
            Dim i As Integer

            myConn.Open()
            For i = 0 To BatchListData.Tables("BatchData").Rows.Count - 1
                BatchListData.Tables("BatchData").Rows(i)("DateCode") = Mid(BatchListData.Tables("BatchData").Rows(i)("DateCode"), 1, 20)
                BatchListData.Tables("BatchData").Rows(i)("LotNo") = Mid(BatchListData.Tables("BatchData").Rows(i)("LotNo"), 1, 20)

                If BatchListData.Tables("BatchData").Rows(i)("ExpiredDate") = "" Then
                    strCMD = String.Format("INSERT INTO T_InvMaster (InvID,Line,Item,Revision,SLOC,Subinventory,Locator,UnitQty,NoofPackage,Manufacturer,ManufacturerPN, CreatedOn, CreatedBy, DateCode, LotNo, BPCSPN) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',getDate(),'{11}','{12}','{13}','{14}')", BatchListData.Tables("BatchData").Rows(i)("InvID"), BatchListData.Tables("BatchData").Rows(i)("Line"), BatchListData.Tables("BatchData").Rows(i)("Item"), "", BatchListData.Tables("BatchData").Rows(i)("SLOC"), BatchListData.Tables("BatchData").Rows(i)("Subinventory"), BatchListData.Tables("BatchData").Rows(i)("Locator"), BatchListData.Tables("BatchData").Rows(i)("UnitQty"), BatchListData.Tables("BatchData").Rows(i)("NoofPackage"), BatchListData.Tables("BatchData").Rows(i)("Manufacturer"), BatchListData.Tables("BatchData").Rows(i)("ManufacturerPN"), OracleLoginData.User.ToUpper, BatchListData.Tables("BatchData").Rows(i)("DateCode"), BatchListData.Tables("BatchData").Rows(i)("LotNo"), "")
                Else
                    strCMD = String.Format("INSERT INTO T_InvMaster (InvID,Line,Item,Revision,SLOC,Subinventory,Locator,ExpiredDate,UnitQty,NoofPackage,Manufacturer,ManufacturerPN, CreatedOn, CreatedBy, DateCode, LotNo, BPCSPN) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}',getDate(),'{12}','{13}','{14}','{15}')", BatchListData.Tables("BatchData").Rows(i)("InvID"), BatchListData.Tables("BatchData").Rows(i)("Line"), BatchListData.Tables("BatchData").Rows(i)("Item"), "", BatchListData.Tables("BatchData").Rows(i)("SLOC"), BatchListData.Tables("BatchData").Rows(i)("Subinventory"), BatchListData.Tables("BatchData").Rows(i)("Locator"), BatchListData.Tables("BatchData").Rows(i)("ExpiredDate"), BatchListData.Tables("BatchData").Rows(i)("UnitQty"), BatchListData.Tables("BatchData").Rows(i)("NoofPackage"), BatchListData.Tables("BatchData").Rows(i)("Manufacturer"), BatchListData.Tables("BatchData").Rows(i)("ManufacturerPN"), OracleLoginData.User.ToUpper, BatchListData.Tables("BatchData").Rows(i)("DateCode"), BatchListData.Tables("BatchData").Rows(i)("LotNo"), "")
                End If
                InvMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                ra = InvMasterSQLCommand.ExecuteNonQuery()
            Next
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("ZSMigration-AddInvMaster_ForBatch", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Private Function CLIDforBatch(ByVal BatchListData As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim i As Integer
            Dim NewCLIDCommand As SqlClient.SqlCommand
            Dim InvLabelSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim rb As Integer
            Dim strCLID As String
            Dim NextCLID As String
            Dim myDataRow As Data.DataRow

            myConn.Open()
            For i = 0 To BatchListData.Tables("CLIDData").Rows.Count - 1
                If FixNull(BatchListData.Tables("CLIDData").Rows(i)("ExpDate")) <> "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocYear,RecDocNo,RecDocItem,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocYear,@RecDocNo,@RecDocItem,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ExpDate,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageType,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                    NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
                    NewCLIDCommand.Parameters("@ExpDate").Value = BatchListData.Tables("CLIDData").Rows(i)("ExpDate")
                ElseIf FixNull(BatchListData.Tables("CLIDData").Rows(i)("ExpDate")) = "" Then
                    NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocYear,RecDocNo,RecDocItem,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProdDate,RoHS,PurOrdNo,PurOrdItem,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType,MaterialDesc,VendorName,VendorPN,SLOC,StorageType,StorageBin,Operator,IsTraceable,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocYear,@RecDocNo,@RecDocItem,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType,@MaterialDesc,@VendorName,@VendorPN,@SLOC,@StorageType,@StorageBin,@Operator,@IsTraceable,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
                End If
                NewCLIDCommand.Parameters.Add("@NewCLID", SqlDbType.VarChar, 20, "CLID")
                NewCLIDCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20, "OrgCode")
                NewCLIDCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 50, "MaterialNo")
                NewCLIDCommand.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 20, "MaterialRevision")
                NewCLIDCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
                NewCLIDCommand.Parameters.Add("@UOM", SqlDbType.VarChar, 10, "UOM")
                NewCLIDCommand.Parameters.Add("@QtyBaseUOM", SqlDbType.Decimal, 13, "QtyBaseUOM")
                NewCLIDCommand.Parameters.Add("@BaseUOM", SqlDbType.VarChar, 10, "BaseUOM")
                NewCLIDCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 20, "DateCode")
                NewCLIDCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 20, "LotNo")
                NewCLIDCommand.Parameters.Add("@COO", SqlDbType.VarChar, 20, "COO")
                NewCLIDCommand.Parameters.Add("@RecDocYear", SqlDbType.VarChar, 50, "RecDocYear")
                NewCLIDCommand.Parameters.Add("@RecDocNo", SqlDbType.VarChar, 50, "RecDocNo")
                NewCLIDCommand.Parameters.Add("@RecDocItem", SqlDbType.VarChar, 10, "RecDocItem")
                'NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
                NewCLIDCommand.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
                NewCLIDCommand.Parameters.Add("@Printed", SqlDbType.VarChar, 100, "Printed")
                NewCLIDCommand.Parameters.Add("@VendorID", SqlDbType.VarChar, 10, "VendorID")
                NewCLIDCommand.Parameters.Add("@RecDate", SqlDbType.SmallDateTime, 4, "RecDate")
                NewCLIDCommand.Parameters.Add("@RoHS", SqlDbType.VarChar, 10, "RoHS")
                NewCLIDCommand.Parameters.Add("@PurOrdNo", SqlDbType.VarChar, 20, "PurOrdNo")
                NewCLIDCommand.Parameters.Add("@PurOrdItem", SqlDbType.VarChar, 10, "PurOrdItem")
                NewCLIDCommand.Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 25, "InvoiceNo")
                NewCLIDCommand.Parameters.Add("@BillofLading", SqlDbType.VarChar, 25, "BillofLading")
                NewCLIDCommand.Parameters.Add("@DN", SqlDbType.VarChar, 25, "DN")
                NewCLIDCommand.Parameters.Add("@HeaderText", SqlDbType.VarChar, 25, "HeaderText")
                'NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
                NewCLIDCommand.Parameters.Add("@ProdDate", SqlDbType.SmallDateTime, 4, "ProdDate")
                NewCLIDCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
                NewCLIDCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 10, "StockType")
                NewCLIDCommand.Parameters.Add("@MaterialDesc", SqlDbType.VarChar, 50, "MaterialDesc")
                NewCLIDCommand.Parameters.Add("@VendorName", SqlDbType.VarChar, 50, "VendorName")
                NewCLIDCommand.Parameters.Add("@VendorPN", SqlDbType.VarChar, 50, "VendorPN")
                NewCLIDCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
                NewCLIDCommand.Parameters.Add("@StorageType", SqlDbType.VarChar, 20, "StorageType")
                NewCLIDCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
                NewCLIDCommand.Parameters.Add("@Operator", SqlDbType.VarChar, 50, "Operator")
                NewCLIDCommand.Parameters.Add("@IsTraceable", SqlDbType.VarChar, 10, "IsTraceable")
                NewCLIDCommand.Parameters.Add("@MatDocNo", SqlDbType.VarChar, 50, "MatDocNo")
                NewCLIDCommand.Parameters.Add("@MatDocItem", SqlDbType.VarChar, 10, "MatDocItem")
                NewCLIDCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 50, "Manufacturer")
                NewCLIDCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 50, "ManufacturerPN")
                NewCLIDCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
                NewCLIDCommand.Parameters.Add("@AddlData", SqlDbType.VarChar, 20, "AddlData")
                NewCLIDCommand.Parameters.Add("@Stemp", SqlDbType.VarChar, 50, "Stemp")
                NewCLIDCommand.Parameters.Add("@MSL", SqlDbType.VarChar, 50, "MSL")
                NewCLIDCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 50, "LastTransaction")

                NextCLID = GetNextCLID(OracleLoginData)
                NewCLIDCommand.Parameters("@NewCLID").Value = NextCLID
                BatchListData.Tables("CLIDData").Rows(i)("NewCLID") = NextCLID
                NewCLIDCommand.Parameters("@OrgCode").Value = BatchListData.Tables("CLIDData").Rows(i)("OrgCode")
                NewCLIDCommand.Parameters("@MaterialNo").Value = BatchListData.Tables("CLIDData").Rows(i)("MaterialNo")
                NewCLIDCommand.Parameters("@MaterialRevision").Value = BatchListData.Tables("CLIDData").Rows(i)("MaterialRevision")
                NewCLIDCommand.Parameters("@Qty").Value = BatchListData.Tables("CLIDData").Rows(i)("Qty")
                NewCLIDCommand.Parameters("@UOM").Value = BatchListData.Tables("CLIDData").Rows(i)("UOM")
                NewCLIDCommand.Parameters("@QtyBaseUOM").Value = BatchListData.Tables("CLIDData").Rows(i)("QtyBaseUOM")
                NewCLIDCommand.Parameters("@BaseUOM").Value = BatchListData.Tables("CLIDData").Rows(i)("BaseUOM")
                NewCLIDCommand.Parameters("@DateCode").Value = BatchListData.Tables("CLIDData").Rows(i)("DateCode")
                NewCLIDCommand.Parameters("@LotNo").Value = BatchListData.Tables("CLIDData").Rows(i)("LotNo")
                NewCLIDCommand.Parameters("@COO").Value = BatchListData.Tables("CLIDData").Rows(i)("COO")
                NewCLIDCommand.Parameters("@RecDocYear").Value = BatchListData.Tables("CLIDData").Rows(i)("RecDocYear")
                NewCLIDCommand.Parameters("@RecDocNo").Value = BatchListData.Tables("CLIDData").Rows(i)("RecDocNo")
                NewCLIDCommand.Parameters("@RecDocItem").Value = BatchListData.Tables("CLIDData").Rows(i)("RecDocItem")
                NewCLIDCommand.Parameters("@CreatedBy").Value = BatchListData.Tables("CLIDData").Rows(i)("CreatedBy")
                NewCLIDCommand.Parameters("@Printed").Value = BatchListData.Tables("CLIDData").Rows(i)("Printed")
                NewCLIDCommand.Parameters("@VendorID").Value = BatchListData.Tables("CLIDData").Rows(i)("VendorID")
                NewCLIDCommand.Parameters("@RecDate").Value = BatchListData.Tables("CLIDData").Rows(i)("RecDate")
                NewCLIDCommand.Parameters("@RoHS").Value = BatchListData.Tables("CLIDData").Rows(i)("RoHS")
                NewCLIDCommand.Parameters("@PurOrdNo").Value = BatchListData.Tables("CLIDData").Rows(i)("PurOrdNo")
                NewCLIDCommand.Parameters("@PurOrdItem").Value = BatchListData.Tables("CLIDData").Rows(i)("PurOrdItem")
                NewCLIDCommand.Parameters("@InvoiceNo").Value = BatchListData.Tables("CLIDData").Rows(i)("InvoiceNo")
                NewCLIDCommand.Parameters("@BillofLading").Value = BatchListData.Tables("CLIDData").Rows(i)("BillofLading")
                NewCLIDCommand.Parameters("@DN").Value = BatchListData.Tables("CLIDData").Rows(i)("DN")
                NewCLIDCommand.Parameters("@HeaderText").Value = BatchListData.Tables("CLIDData").Rows(i)("HeaderText")
                'NewCLIDCommand.Parameters("@ExpDate").Value = BatchListData.Tables("CLIDData").Rows(i)("ExpDate")
                NewCLIDCommand.Parameters("@ProdDate").Value = BatchListData.Tables("CLIDData").Rows(i)("ProdDate")
                NewCLIDCommand.Parameters("@ReasonCode").Value = BatchListData.Tables("CLIDData").Rows(i)("ReasonCode")
                NewCLIDCommand.Parameters("@StockType").Value = BatchListData.Tables("CLIDData").Rows(i)("StockType")
                NewCLIDCommand.Parameters("@MaterialDesc").Value = SQLString(BatchListData.Tables("CLIDData").Rows(i)("MaterialDesc"))
                NewCLIDCommand.Parameters("@VendorName").Value = BatchListData.Tables("CLIDData").Rows(i)("VendorName")
                NewCLIDCommand.Parameters("@VendorPN").Value = BatchListData.Tables("CLIDData").Rows(i)("VendorPN")
                NewCLIDCommand.Parameters("@SLOC").Value = BatchListData.Tables("CLIDData").Rows(i)("SLOC")
                NewCLIDCommand.Parameters("@StorageType").Value = BatchListData.Tables("CLIDData").Rows(i)("StorageType")
                NewCLIDCommand.Parameters("@StorageBin").Value = BatchListData.Tables("CLIDData").Rows(i)("StorageBin")
                NewCLIDCommand.Parameters("@Operator").Value = BatchListData.Tables("CLIDData").Rows(i)("Operator")
                NewCLIDCommand.Parameters("@IsTraceable").Value = BatchListData.Tables("CLIDData").Rows(i)("IsTraceable")
                NewCLIDCommand.Parameters("@MatDocNo").Value = BatchListData.Tables("CLIDData").Rows(i)("MatDocNo")
                NewCLIDCommand.Parameters("@MatDocItem").Value = BatchListData.Tables("CLIDData").Rows(i)("MatDocItem")
                NewCLIDCommand.Parameters("@Manufacturer").Value = BatchListData.Tables("CLIDData").Rows(i)("Manufacturer")
                NewCLIDCommand.Parameters("@ManufacturerPN").Value = BatchListData.Tables("CLIDData").Rows(i)("ManufacturerPN")
                NewCLIDCommand.Parameters("@QMLStatus").Value = BatchListData.Tables("CLIDData").Rows(i)("QMLStatus")
                NewCLIDCommand.Parameters("@AddlData").Value = BatchListData.Tables("CLIDData").Rows(i)("AddlData")
                NewCLIDCommand.Parameters("@Stemp").Value = BatchListData.Tables("CLIDData").Rows(i)("Stemp")
                NewCLIDCommand.Parameters("@MSL").Value = BatchListData.Tables("CLIDData").Rows(i)("MSL")
                NewCLIDCommand.Parameters("@LastTransaction").Value = "ZSMigration"

                NewCLIDCommand.CommandType = CommandType.Text
                ra = NewCLIDCommand.ExecuteNonQuery()

                myDataRow = BatchListData.Tables("CLIDTable").NewRow()
                myDataRow("InvID") = BatchListData.Tables("CLIDData").Rows(i)("InvID")
                myDataRow("Line") = BatchListData.Tables("CLIDData").Rows(i)("Line")
                myDataRow("CLID") = NextCLID
                BatchListData.Tables("CLIDTable").Rows.Add(myDataRow)

                strCLID = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", BatchListData.Tables("CLIDData").Rows(i)("InvID"), BatchListData.Tables("CLIDData").Rows(i)("Line"), NextCLID)
                InvLabelSQLCommand = New SqlClient.SqlCommand(strCLID, myConn)
                rb = InvLabelSQLCommand.ExecuteNonQuery()

            Next
            myConn.Close()
            Return BatchListData
        Catch ex As Exception
            ErrorLogging("ZSMigration-CLIDforBatch", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Private Function PRODforBatch(ByVal BatchListData As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            Dim myDataRow As Data.DataRow
            Dim CLMasterSQLCommand As SqlClient.SqlCommand
            Dim InvLabelSQLCommand As SqlClient.SqlCommand
            Dim ra As Integer
            Dim rb As Integer
            Dim strCMD As String
            Dim strProd As String
            Dim ProdCLID, OrgCode, matlNo, matlDesc, UOM, PurOrdNo, statusCode, RecDocYear, RecDocNo, ExpDate, DateCode, LotNo, Revision, CreatedBy, SLOC, StorageType, StorageBin, Manufacturer, ManufacturerPN As String
            Dim QtyBaseUOM As Decimal
            Dim i As Integer
            Dim result1 As New Object
            Dim Last_Transaction As String

            myConn.Open()
            For i = 0 To BatchListData.Tables("ProdData").Rows.Count - 1
                ProdCLID = GetNextProdID(OracleLoginData)
                BatchListData.Tables("ProdData").Rows(i)("CLID") = ProdCLID
                OrgCode = BatchListData.Tables("ProdData").Rows(i)("OrgCode")
                matlNo = BatchListData.Tables("ProdData").Rows(i)("MaterialNo")
                matlDesc = BatchListData.Tables("ProdData").Rows(i)("MaterialDesc")
                UOM = BatchListData.Tables("ProdData").Rows(i)("BaseUOM")
                PurOrdNo = BatchListData.Tables("ProdData").Rows(i)("PurOrdNo")
                QtyBaseUOM = BatchListData.Tables("ProdData").Rows(i)("QtyBaseUOM")
                RecDocYear = BatchListData.Tables("ProdData").Rows(i)("RecDocYear")
                RecDocNo = BatchListData.Tables("ProdData").Rows(i)("RecDocNo")
                DateCode = BatchListData.Tables("ProdData").Rows(i)("DateCode")
                LotNo = BatchListData.Tables("ProdData").Rows(i)("LotNo")
                statusCode = "1"
                CreatedBy = OracleLoginData.User.ToUpper
                Revision = BatchListData.Tables("ProdData").Rows(i)("MaterialRevision")
                SLOC = BatchListData.Tables("ProdData").Rows(i)("SLOC")
                StorageType = BatchListData.Tables("ProdData").Rows(i)("StorageType")
                StorageBin = BatchListData.Tables("ProdData").Rows(i)("StorageBin")
                Manufacturer = BatchListData.Tables("ProdData").Rows(i)("Manufacturer")
                ManufacturerPN = BatchListData.Tables("ProdData").Rows(i)("ManufacturerPN")
                Last_Transaction = "ZSMigration"

                If FixNull(BatchListData.Tables("ProdData").Rows(i)("ExpDate")) = "" Then
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,MaterialDesc,Qty,UOM,QtyBaseUOM,BaseUOM,CreatedOn,CreatedBy,DateCode,LotNo,statusCode,RecDocYear,RecDocNo,RecDate,PurOrdNo,SLOC,StorageType,StorageBin,Manufacturer,ManufacturerPN,LastTransaction) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',getDate(),'{9}','{10}','{11}','{12}','{13}','{14}',getDate(),'{15}','{16}','{17}','{18}','{19}','{20}','{21}')", ProdCLID, OrgCode, matlNo, Revision, matlDesc, QtyBaseUOM, UOM, QtyBaseUOM, UOM, CreatedBy, DateCode, LotNo, statusCode, RecDocYear, RecDocNo, PurOrdNo, SLOC, StorageType, StorageBin, Manufacturer, ManufacturerPN, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", BatchListData.Tables("ProdData").Rows(i)("InvID"), BatchListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                ElseIf FixNull(BatchListData.Tables("ProdData").Rows(i)("ExpDate")) <> "" Then
                    ExpDate = BatchListData.Tables("ProdData").Rows(i)("ExpDate")
                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,MaterialDesc,Qty,UOM,QtyBaseUOM,BaseUOM,CreatedOn,CreatedBy,DateCode,LotNo,statusCode,ExpDate,RecDocYear,RecDocNo,RecDate,PurOrdNo,SLOC,StorageType,StorageBin,Manufacturer,ManufacturerPN,LastTransaction) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}',getDate(),'{16}','{17}','{18}','{19}','{20}','{21}','{22}')", ProdCLID, OrgCode, matlNo, Revision, matlDesc, QtyBaseUOM, UOM, QtyBaseUOM, UOM, CreatedBy, DateCode, LotNo, statusCode, ExpDate, RecDocYear, RecDocNo, PurOrdNo, SLOC, StorageType, StorageBin, Manufacturer, ManufacturerPN, Last_Transaction)
                    strProd = String.Format("INSERT INTO T_InvLabel (InvID,Line,CLID) values ('{0}','{1}','{2}')", BatchListData.Tables("ProdData").Rows(i)("InvID"), BatchListData.Tables("ProdData").Rows(i)("Line"), ProdCLID)
                End If

                'myConn.Open()
                CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                InvLabelSQLCommand = New SqlClient.SqlCommand(strProd, myConn)
                ra = CLMasterSQLCommand.ExecuteNonQuery()
                rb = InvLabelSQLCommand.ExecuteNonQuery()
                'myConn.Close()

                myDataRow = BatchListData.Tables("CLIDTable").NewRow()
                myDataRow("InvID") = BatchListData.Tables("ProdData").Rows(i)("InvID")
                myDataRow("Line") = BatchListData.Tables("ProdData").Rows(i)("Line")
                myDataRow("CLID") = ProdCLID
                BatchListData.Tables("CLIDTable").Rows.Add(myDataRow)
            Next
            myConn.Close()
            Return BatchListData
        Catch ex As Exception
            ErrorLogging("ZSMigration-ProdIDforBatch", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function BJ_GetDJInfo(ByVal DJ As String, ByVal PWC As String, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                'ErrorLogging("MaterialMovement-BJ_GetDJInfo", OracleLoginData.User, "DJ: " & DJ & " OrgID: " & OracleLoginData.OrgID)
                OracleLoginData.OrgID = GetOrgID(OracleLoginData.OrgCode)

                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_invcc_pkg.get_pre_work_data"              'Get IR / ISO

                oda.SelectCommand.Parameters.Add("p_org_id", OracleType.Int32).Value = OracleLoginData.OrgID
                oda.SelectCommand.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                oda.SelectCommand.Parameters.Add("p_pwc", OracleType.VarChar, 50).Value = PWC
                oda.SelectCommand.Parameters.Add("o_err_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_bj_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_bj_lines", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If FixNull(oda.SelectCommand.Parameters("o_err_msg").Value) <> "" Then
                    BJ_GetDJInfo.BJInfo = Nothing
                    BJ_GetDJInfo.Flag = oda.SelectCommand.Parameters("o_err_flag").Value
                    BJ_GetDJInfo.ErrMsg = oda.SelectCommand.Parameters("o_err_msg").Value
                Else
                    ds.Tables(0).TableName = "dj_header"
                    ds.Tables(1).TableName = "dj_lines"
                    BJ_GetDJInfo.BJInfo = ds
                    BJ_GetDJInfo.Flag = "Y"
                    BJ_GetDJInfo.ErrMsg = ""
                End If
                Return BJ_GetDJInfo
            Catch oe As Exception
                ErrorLogging("MaterialMovement-BJ_GetDJInfo", OracleLoginData.User, oe.Message & oe.Source, "E")
                BJ_GetDJInfo.BJInfo = Nothing
                BJ_GetDJInfo.Flag = "N"
                BJ_GetDJInfo.ErrMsg = oe.Message & oe.Source
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetNextBJID(ByVal OracleLoginData As ERPLogin) As String
        Dim NextInvID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try
            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "BJID")

            'Try up to 5 times when failed getting next id
            NextInvID = ""
            k = 0
            While (k < 5 And NextInvID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextInvID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextBJID", "Deadlocked? " & Str(k), "Failed to get next BJID, " & ex.Message & ex.Source)
                End Try
            End While
            myConn.Close()
            Return NextInvID
        Catch ex As Exception
            ErrorLogging("GetNextBJID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextInvID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function BJ_Creation(ByVal DJInfo As DataSet, ByVal BJds As DataSet, ByVal PWC As String, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Using da As DataAccess = GetDataAccess()
            Dim BJ, tmp_bpa, sqlstr As String
            Dim ProcessCode() As String
            Dim i As Integer
            Dim dr, dr2 As DataRow
            Dim dr_del() As DataRow = Nothing

            Try
                'Delete records which ProcessCode is blank or without '-'
                dr_del = DJInfo.Tables("dj_lines").Select("(process_code is null or process_code = '' or process_code not like '%-%')")
                If dr_del.Length > 0 Then
                    For Each drdel As DataRow In dr_del
                        DJInfo.Tables("dj_lines").Rows.Remove(drdel)
                    Next
                End If

                If DJInfo.Tables("dj_lines").Rows.Count < 1 Then
                    BJ_Creation.BJInfo = Nothing
                    BJ_Creation.Flag = "1"
                    BJ_Creation.ErrMsg = "No component with ProcessCode!"
                    Return BJ_Creation
                End If

                'Sort Process_Code by Ascending before creating BJ
                Dim DJ_Lines As DataTable = New DataTable("dj_lines")
                Dim SortColName As String = DJInfo.Tables("dj_lines").Columns(2).ColumnName
                SortColName = SortColName & " ASC"
                DJInfo.Tables("dj_lines").DefaultView.Sort = SortColName
                DJ_Lines = DJInfo.Tables("dj_lines").DefaultView.ToTable()
                DJInfo.Tables.Remove("dj_lines")
                DJInfo.Tables.Add(DJ_Lines)

                'Create BJ
                For i = 0 To DJInfo.Tables("dj_lines").Rows.Count - 1
                    If i = 0 Then
                        Erase ProcessCode
                        ProcessCode = Split(DJInfo.Tables("dj_lines").Rows(i)("process_code"), "-")
                        tmp_bpa = ProcessCode(0)
                        BJ = GetNextBJID(OracleLoginData)
                        'BJ Header
                        dr = BJds.Tables("BJHeader").NewRow
                        dr("ProdLine") = DJInfo.Tables("dj_header").Rows(0)("production_line")
                        dr("DJ") = DJInfo.Tables("dj_header").Rows(0)("dj")
                        dr("TLA") = DJInfo.Tables("dj_header").Rows(0)("model")
                        dr("BJ") = BJ
                        dr("BPA") = tmp_bpa
                        dr("BJQty") = DJInfo.Tables("dj_header").Rows(0)("start_quantity")
                        dr("CmpQty") = 0
                        dr("CurCmpQty") = 0
                        dr("Status") = "Released"
                        dr("SchDate") = DateAdd(DateInterval.Hour, -8, DateAdd(DateInterval.Day, 2, Date.Now.Date))  'Date.Now.Date   'DateAdd("d", 1, Today())
                        dr("Remarks") = ""

                        dr("DJReleasedDate") = DJInfo.Tables("dj_header").Rows(0)("DJ_Released_Date")
                        dr("PWC") = PWC
                        dr("PWCSubInv") = DJInfo.Tables("dj_header").Rows(0)("PWC_SubInventory")
                        dr("PWCLocator") = DJInfo.Tables("dj_header").Rows(0)("PWC_Locator")
                        dr("TLASubInv") = DJInfo.Tables("dj_header").Rows(0)("TLA_SubInventory")
                        dr("TLALocator") = DJInfo.Tables("dj_header").Rows(0)("TLA_Locator")
                        dr("CreatedOn") = DateTime.Now
                        dr("CreatedBy") = OracleLoginData.User
                        dr("ChangedOn") = DBNull.Value
                        dr("ChangedBy") = ""
                        dr("OpenQty") = dr("BJQty")
                        BJds.Tables("BJHeader").Rows.Add(dr)

                        'sqlstr = String.Format("Insert into T_PWCBJHeader (BJ, DJ, DJQty, TLA, ProdLine, PWC, BPA, BJQty, OpenQty, Status, ScheduleDate, CreatedOn, CreatedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')", dr("BJ"), dr("DJ"), dr("BJQty"), dr("TLA"), dr("ProdLine"), dr("PWC"), dr("BPA"), dr("BJQty"), dr("OpenQty"), dr("Status"), dr("SchDate"), dr("CreatedOn"), dr("CreatedBy"), dr("Remarks"))
                        'da.ExecuteNonQuery(sqlstr)

                        'BJ Item
                        dr2 = BJds.Tables("BJItem").NewRow
                        dr2("Item") = DJInfo.Tables("dj_lines").Rows(0)("component_item")
                        dr2("ProcessCode") = DJInfo.Tables("dj_lines").Rows(0)("Process_Code")
                        dr2("Usage") = DJInfo.Tables("dj_lines").Rows(0)("cmp_usage")
                        dr2("ReqQty") = DJInfo.Tables("dj_lines").Rows(0)("usage") * dr("BJQty")
                        If DJInfo.Tables("dj_lines").Rows(0)("ejit_flag") = "Y" OrElse DJInfo.Tables("dj_lines").Rows(0)("kanban_flag") = "Y" Then
                            dr2("SupplyType") = "eJIT"
                        ElseIf DJInfo.Tables("dj_lines").Rows(0)("consumables_flag") = "Y" Then
                            dr2("SupplyType") = "Consumable"
                        Else
                            dr2("SupplyType") = "Normal"
                        End If

                        dr2("MoveOrder") = ""
                        dr2("PickedQty") = 0
                        dr2("Remarks") = ""

                        dr2("BJ") = BJ
                        dr2("ChangedOn") = DateTime.Now
                        dr2("ChangedBy") = OracleLoginData.User
                        BJds.Tables("BJItem").Rows.Add(dr2)

                        'sqlstr = String.Format("Insert into T_PWCBJItem (BJ, Component, ProcessCode, Usage, SupplyType, MO, ChangedOn, ChangedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", dr("BJ"), dr("Item"), dr("ProcessCode"), dr("Usage"), dr("SupplyType"), dr("MoveOrder"), dr("ChangedOn"), dr("ChangedBy"), dr("Remarks"))
                        'da.ExecuteNonQuery(sqlstr)
                    Else
                        Erase ProcessCode
                        ProcessCode = Split(DJInfo.Tables("dj_lines").Rows(i)("process_code"), "-")
                        If tmp_bpa <> ProcessCode(0) Then
                            tmp_bpa = ProcessCode(0)
                            BJ = GetNextBJID(OracleLoginData)

                            'BJ Header
                            dr = BJds.Tables("BJHeader").NewRow
                            dr("ProdLine") = DJInfo.Tables("dj_header").Rows(0)("production_line")
                            dr("DJ") = DJInfo.Tables("dj_header").Rows(0)("dj")
                            dr("TLA") = DJInfo.Tables("dj_header").Rows(0)("model")
                            dr("BJ") = BJ
                            dr("BPA") = tmp_bpa
                            dr("BJQty") = DJInfo.Tables("dj_header").Rows(0)("start_quantity")
                            dr("CmpQty") = 0
                            dr("CurCmpQty") = 0
                            dr("Status") = "Released"
                            dr("SchDate") = DateAdd(DateInterval.Hour, -8, DateAdd(DateInterval.Day, 2, Date.Now.Date))  'Date.Now.Date   'DateAdd("d", 1, Today())
                            dr("Remarks") = ""

                            dr("DJReleasedDate") = DJInfo.Tables("dj_header").Rows(0)("DJ_Released_Date")
                            dr("PWC") = PWC
                            dr("PWCSubInv") = DJInfo.Tables("dj_header").Rows(0)("PWC_SubInventory")
                            dr("PWCLocator") = DJInfo.Tables("dj_header").Rows(0)("PWC_Locator")
                            dr("TLASubInv") = DJInfo.Tables("dj_header").Rows(0)("TLA_SubInventory")
                            dr("TLALocator") = DJInfo.Tables("dj_header").Rows(0)("TLA_Locator")
                            dr("CreatedOn") = DateTime.Now
                            dr("CreatedBy") = OracleLoginData.User
                            dr("ChangedOn") = DBNull.Value
                            dr("ChangedBy") = ""
                            dr("OpenQty") = dr("BJQty")
                            BJds.Tables("BJHeader").Rows.Add(dr)

                            'sqlstr = String.Format("Insert into T_PWCBJHeader (BJ, DJ, DJQty, TLA, ProdLine, PWC, BPA, BJQty, OpenQty, Status, ScheduleDate, CreatedOn, CreatedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')", dr("BJ"), dr("DJ"), dr("BJQty"), dr("TLA"), dr("ProdLine"), dr("PWC"), dr("BPA"), dr("BJQty"), dr("OpenQty"), dr("Status"), dr("SchDate"), dr("CreatedOn"), dr("CreatedBy"), dr("Remarks"))
                            'da.ExecuteNonQuery(sqlstr)

                            'BJ Item
                            dr2 = BJds.Tables("BJItem").NewRow
                            dr2("Item") = DJInfo.Tables("dj_lines").Rows(i)("component_item")
                            dr2("ProcessCode") = DJInfo.Tables("dj_lines").Rows(i)("Process_Code")
                            dr2("Usage") = DJInfo.Tables("dj_lines").Rows(i)("cmp_usage")
                            dr2("ReqQty") = DJInfo.Tables("dj_lines").Rows(i)("usage") * dr("BJQty")
                            If DJInfo.Tables("dj_lines").Rows(i)("ejit_flag") = "Y" OrElse DJInfo.Tables("dj_lines").Rows(i)("kanban_flag") = "Y" Then
                                dr2("SupplyType") = "eJIT"
                            ElseIf DJInfo.Tables("dj_lines").Rows(i)("consumables_flag") = "Y" Then
                                dr2("SupplyType") = "Consumable"
                            Else
                                dr2("SupplyType") = "Normal"
                            End If

                            dr2("MoveOrder") = ""
                            dr2("PickedQty") = 0
                            dr2("Remarks") = ""

                            dr2("BJ") = BJ
                            dr2("ChangedOn") = DateTime.Now
                            dr2("ChangedBy") = OracleLoginData.User
                            BJds.Tables("BJItem").Rows.Add(dr2)

                            'sqlstr = String.Format("Insert into T_PWCBJItem (BJ, Component, ProcessCode, Usage, SupplyType, MO, ChangedOn, ChangedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", dr("BJ"), dr("Item"), dr("ProcessCode"), dr("Usage"), dr("SupplyType"), dr("MoveOrder"), dr("ChangedOn"), dr("ChangedBy"), dr("Remarks"))
                            'da.ExecuteNonQuery(sqlstr)
                        Else
                            'BJ Item
                            dr2 = BJds.Tables("BJItem").NewRow
                            dr2("Item") = DJInfo.Tables("dj_lines").Rows(i)("component_item")
                            dr2("ProcessCode") = DJInfo.Tables("dj_lines").Rows(i)("Process_Code")
                            dr2("Usage") = DJInfo.Tables("dj_lines").Rows(i)("cmp_usage")
                            dr2("ReqQty") = DJInfo.Tables("dj_lines").Rows(i)("usage") * dr("BJQty")
                            If DJInfo.Tables("dj_lines").Rows(i)("ejit_flag") = "Y" OrElse DJInfo.Tables("dj_lines").Rows(i)("kanban_flag") = "Y" Then
                                dr2("SupplyType") = "eJIT"
                            ElseIf DJInfo.Tables("dj_lines").Rows(i)("consumables_flag") = "Y" Then
                                dr2("SupplyType") = "Consumable"
                            Else
                                dr2("SupplyType") = "Normal"
                            End If

                            dr2("MoveOrder") = ""
                            dr2("PickedQty") = 0
                            dr2("Remarks") = ""

                            dr2("BJ") = BJ
                            dr2("ChangedOn") = DateTime.Now
                            dr2("ChangedBy") = OracleLoginData.User
                            BJds.Tables("BJItem").Rows.Add(dr2)

                            'sqlstr = String.Format("Insert into T_PWCBJItem (BJ, Component, ProcessCode, Usage, SupplyType, MO, ChangedOn, ChangedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", dr("BJ"), dr("Item"), dr("ProcessCode"), dr("Usage"), dr("SupplyType"), dr("MoveOrder"), dr("ChangedOn"), dr("ChangedBy"), dr("Remarks"))
                            'da.ExecuteNonQuery(sqlstr)
                        End If
                    End If
                Next
                BJ_Creation.BJInfo = BJds
                BJ_Creation.Flag = "0"
                BJ_Creation.ErrMsg = ""
                Return BJ_Creation
            Catch ex As Exception
                ErrorLogging("MaterialMovement-BJ_Creation", OracleLoginData.User, ex.Message & ex.Source, "E")
                BJ_Creation.BJInfo = Nothing
                BJ_Creation.Flag = "2"
                BJ_Creation.ErrMsg = ex.Message & ex.Source
                Return BJ_Creation
            End Try
        End Using
    End Function

    Public Function isValid_ItemLotControl(ByVal org_code As String, ByVal item_num As String) As Integer
        Using da As DataAccess = GetDataAccess()

            Dim sc As New SqlClient.SqlCommand
            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
            Dim lot_control_code As String
            Dim strCMD As String
            Try
                strCMD = String.Format("exec ora_valid_item_lot_control '{0}','{1}'", org_code, item_num)
                lot_control_code = da.ExecuteScalar(strCMD)

                Return lot_control_code
            Catch oe As Exception
                ErrorLogging("isValid_ItemLotControl", "", oe.Message & oe.Source, "E")
                Return -1
                'Finally
                '    If sc.Connection.State <> ConnectionState.Closed Then sc.Connection.Close()
            End Try
        End Using
    End Function

    Public Function BJ_SaveChange(ByVal BJInfo As DataSet, ByVal BJInitial As DataSet, ByVal txtDJ As String, ByVal CreateFlag As Boolean, ByVal OracleLoginData As ERPLogin) As String
        Dim BJ, DJ, BPA, Status, Remarks, ChangedBy, sqlstr As String
        Dim i, j, LotFlag As Integer
        Dim OpenQty As Decimal
        Dim SchDate, ChangedOn As DateTime
        Dim dr As DataRow
        Dim drow() As DataRow = Nothing
        Dim dtrw() As DataRow = Nothing
        Dim drep() As DataRow = Nothing

        Dim ds As New DataSet
        Dim Tb_Transfer As Data.DataTable = New Data.DataTable("transfer_table")

        Dim p_item_segment1 As DataColumn = New DataColumn("p_item_segment1", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_item_segment1)
        Dim p_subinventory_source As DataColumn = New DataColumn("p_subinventory_source", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_subinventory_source)
        Dim p_locator_source As DataColumn = New DataColumn("p_locator_source", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_locator_source)
        Dim p_subinventory_destination As DataColumn = New DataColumn("p_subinventory_destination", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_subinventory_destination)
        Dim p_locator_destination As DataColumn = New DataColumn("p_locator_destination", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_locator_destination)
        Dim p_transaction_quantity As DataColumn = New DataColumn("p_transaction_quantity", System.Type.GetType("System.Decimal"))
        Tb_Transfer.Columns.Add(p_transaction_quantity)
        Dim p_primary_quantity As DataColumn = New DataColumn("p_primary_quantity", System.Type.GetType("System.Decimal"))
        Tb_Transfer.Columns.Add(p_primary_quantity)
        Dim p_lot_number As DataColumn = New DataColumn("p_lot_number", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(p_lot_number)
        Dim o_return_status As DataColumn = New DataColumn("o_return_status", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(o_return_status)
        Dim o_return_message As DataColumn = New DataColumn("o_return_message", System.Type.GetType("System.String"))
        Tb_Transfer.Columns.Add(o_return_message)
        ds.Tables.Add(Tb_Transfer)

        Try
            Using da As DataAccess = GetDataAccess()
                For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                    drow = Nothing
                    drow = BJInitial.Tables("BJHeader").Select("BJ = '" & BJInfo.Tables("BJHeader").Rows(i)("BJ") & "'")
                    If drow.Length > 0 AndAlso BJInfo.Tables("BJHeader").Rows(i)("CmpQty") <> drow(0)("CmpQty") Then
                        dtrw = Nothing
                        dtrw = BJInfo.Tables("BJItem").Select("BJ = '" & BJInfo.Tables("BJHeader").Rows(i)("BJ") & "' and SupplyType = 'Consumable'")
                        For j = 0 To dtrw.Length - 1
                            dr = ds.Tables("transfer_table").NewRow
                            dr("p_item_segment1") = dtrw(j)("item")
                            If BJInfo.Tables("BJHeader").Rows(i)("CmpQty") > drow(0)("CmpQty") Then
                                dr("p_subinventory_source") = BJInfo.Tables("BJHeader").Rows(i)("PWCSubInv")
                                dr("p_locator_source") = BJInfo.Tables("BJHeader").Rows(i)("PWCLocator")
                                dr("p_subinventory_destination") = BJInfo.Tables("BJHeader").Rows(i)("TLASubInv")
                                dr("p_locator_destination") = BJInfo.Tables("BJHeader").Rows(i)("TLALocator")
                                dr("p_transaction_quantity") = (BJInfo.Tables("BJHeader").Rows(i)("CmpQty") - drow(0)("CmpQty")) * dtrw(j)("Usage")
                                dr("p_primary_quantity") = dr("p_transaction_quantity")
                                drep = Nothing
                                drep = ds.Tables("transfer_table").Select("p_item_segment1 = '" & dr("p_item_segment1") & "' and p_subinventory_source = '" & dr("p_subinventory_source") & "' and p_locator_source = '" & dr("p_locator_source") & "' and p_subinventory_destination = '" & dr("p_subinventory_destination") & "' and p_locator_destination = '" & dr("p_locator_destination") & "'")
                                If drep.Length > 0 Then
                                    drep(0)("p_transaction_quantity") = drep(0)("p_transaction_quantity") + dr("p_transaction_quantity")
                                    drep(0)("p_primary_quantity") = drep(0)("p_transaction_quantity")
                                Else
                                    LotFlag = isValid_ItemLotControl(OracleLoginData.OrgCode, dr("p_item_segment1"))
                                    If LotFlag = 2 Then
                                        dr("p_lot_number") = "00000"
                                    Else
                                        dr("p_lot_number") = ""
                                    End If
                                    dr("o_return_status") = ""
                                    dr("o_return_message") = ""
                                    ds.Tables("transfer_table").Rows.Add(dr)
                                End If
                            ElseIf BJInfo.Tables("BJHeader").Rows(i)("CmpQty") < drow(0)("CmpQty") Then
                                dr("p_subinventory_source") = BJInfo.Tables("BJHeader").Rows(i)("TLASubInv")
                                dr("p_locator_source") = BJInfo.Tables("BJHeader").Rows(i)("TLALocator")
                                dr("p_subinventory_destination") = BJInfo.Tables("BJHeader").Rows(i)("PWCSubInv")
                                dr("p_locator_destination") = BJInfo.Tables("BJHeader").Rows(i)("PWCLocator")
                                dr("p_transaction_quantity") = (drow(0)("CmpQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty")) * dtrw(j)("Usage")
                                dr("p_primary_quantity") = dr("p_transaction_quantity")
                                drep = Nothing
                                drep = ds.Tables("transfer_table").Select("p_item_segment1 = '" & dr("p_item_segment1") & "' and p_subinventory_source = '" & dr("p_subinventory_source") & "' and p_locator_source = '" & dr("p_locator_source") & "' and p_subinventory_destination = '" & dr("p_subinventory_destination") & "' and p_locator_destination = '" & dr("p_locator_destination") & "'")
                                If drep.Length > 0 Then
                                    drep(0)("p_transaction_quantity") = drep(0)("p_transaction_quantity") + dr("p_transaction_quantity")
                                    drep(0)("p_primary_quantity") = drep(0)("p_transaction_quantity")
                                Else
                                    LotFlag = isValid_ItemLotControl(OracleLoginData.OrgCode, dr("p_item_segment1"))
                                    If LotFlag = 2 Then
                                        dr("p_lot_number") = "00000"
                                    Else
                                        dr("p_lot_number") = ""
                                    End If
                                    dr("o_return_status") = ""
                                    dr("o_return_message") = ""
                                    ds.Tables("transfer_table").Rows.Add(dr)
                                End If
                            End If
                        Next
                    End If
                Next

                If ds.Tables("transfer_table").Rows.Count > 0 Then
                    For i = 0 To ds.Tables("transfer_table").Rows.Count - 1
                        If ds.Tables("transfer_table").Rows(i).RowState = DataRowState.Unchanged Then
                            ds.Tables("transfer_table").Rows(i).SetAdded()
                        End If
                    Next

                    'Oracle transaction
                    Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
                    Dim comm As OracleCommand = da.Ora_Command_Trans()
                    Dim aa As OracleString
                    Dim resp As Integer = OracleLoginData.RespID_Inv
                    Dim appl As Integer = OracleLoginData.AppID_Inv
                    Dim TransactionID As Long
                    Dim OrgID As String
                    OrgID = GetOrgID(OracleLoginData.OrgCode)
                    TransactionID = CLng(GetNextHeaderID(OracleLoginData))

                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
                    comm.ExecuteOracleNonQuery(aa)
                    comm.Parameters.Clear()

                    comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.subinventory_transfer_batch"
                    comm.Parameters.Add("p_timeout", OracleType.Double).Value = 1000 * 60 * 30
                    comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID
                    comm.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                    'comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
                    comm.Parameters.Add("p_source_line_id", OracleType.Double).Value = 99
                    comm.Parameters.Add("p_source_header_id", OracleType.Double).Value = 99
                    comm.Parameters.Add("p_user_id", OracleType.Double).Value = OracleLoginData.UserID
                    comm.Parameters.Add("p_item_segment1", OracleType.VarChar, 240).SourceColumn = "p_item_segment1"
                    'comm.Parameters.Add("p_item_revision", OracleType.VarChar, 240).SourceColumn = "p_item_revision"
                    comm.Parameters.Add("p_subinventory_source", OracleType.VarChar, 240).SourceColumn = "p_subinventory_source"
                    comm.Parameters.Add("p_locator_source", OracleType.VarChar, 240).SourceColumn = "p_locator_source"
                    comm.Parameters.Add("p_subinventory_destination", OracleType.VarChar, 240).SourceColumn = "p_subinventory_destination"
                    comm.Parameters.Add("p_locator_destination", OracleType.VarChar, 240).SourceColumn = "p_locator_destination"
                    comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240).SourceColumn = "p_lot_number"
                    'comm.Parameters.Add("p_lot_number", OracleType.VarChar, 240).Value = "00000"
                    'comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
                    comm.Parameters.Add("p_transaction_quantity", OracleType.Double).SourceColumn = "p_transaction_quantity"
                    comm.Parameters.Add("p_primary_quantity", OracleType.Double).SourceColumn = "p_primary_quantity"
                    'comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240)
                    'comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
                    comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240).Value = "Subinventory Transfer"
                    comm.Parameters.Add("p_etrace_tran_type", OracleType.VarChar, 240).Value = "BJ Completion/BJ Reversal"
                    comm.Parameters.Add("o_return_status", OracleType.VarChar, 240).SourceColumn = "o_return_status"
                    comm.Parameters.Add("o_return_message", OracleType.VarChar, 240).SourceColumn = "o_return_message"

                    comm.Parameters("p_subinventory_destination").Direction = ParameterDirection.InputOutput
                    comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
                    comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

                    oda_h.InsertCommand = comm
                    oda_h.Update(ds.Tables("transfer_table"))

                    Dim result_flag As String
                    drow = ds.Tables("transfer_table").Select("o_return_status <> 'Y'")

                    If drow.Length = 0 Then
                        result_flag = submit_inv_cvs("BJ_SaveChange", TransactionID, OracleLoginData)
                        If result_flag = "Y" Then

                        Else
                            ErrorLogging("BJ_SaveChange", OracleLoginData.User, result_flag, "I")
                            BJ_SaveChange = result_flag
                            result_flag = del_inv_cvs("BJ_SaveChange", TransactionID, OracleLoginData)
                            Return BJ_SaveChange
                        End If
                    Else
                        'For i = 0 To (drow.Length - 1)
                        '    result_flag = result_flag & drow(i)("o_return_message").ToString & " "
                        '    ErrorLogging("BJ_SaveChange", OracleLoginData.User, drow(i)("o_return_message").ToString, "I")
                        'Next
                        result_flag = drow(0)("o_return_message").ToString
                        ErrorLogging("BJ_SaveChange", OracleLoginData.User, drow(0)("o_return_message").ToString, "I")

                        BJ_SaveChange = result_flag  ' collect all error message as result
                        result_flag = del_inv_cvs("BJ_SaveChange", TransactionID, OracleLoginData)
                        Return BJ_SaveChange
                    End If
                End If

                'eTrace transaction
                If CreateFlag = True Then
                    If txtDJ <> "" Then
                        sqlstr = String.Format("DELETE FROM T_PWCBJItem FROM T_PWCBJItem INNER JOIN T_PWCBJHeader ON T_PWCBJItem.BJ = T_PWCBJHeader.BJ WHERE (T_PWCBJHeader.OrgCode = '{0}') AND (T_PWCBJHeader.DJ = '{1}') AND (T_PWCBJHeader.BJQty = T_PWCBJHeader.OpenQty) AND (T_PWCBJHeader.Status = 'Cancelled')", OracleLoginData.OrgCode, txtDJ)
                        da.ExecuteNonQuery(sqlstr)
                        sqlstr = String.Format("DELETE FROM T_PWCBJHeader WHERE OrgCode = '{0}' AND DJ = '{1}' AND BJQty = OpenQty AND Status = 'Cancelled'", OracleLoginData.OrgCode, txtDJ)
                        da.ExecuteNonQuery(sqlstr)

                        sqlstr = String.Format("select top (1) BJ from T_PWCBJHeader where OrgCode = '{0}' AND DJ = '{1}'", OracleLoginData.OrgCode, txtDJ)
                        BJ = FixNull(da.ExecuteScalar(sqlstr))
                        If BJ = "" Then
                            'Create BJ
                            For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                                BJInfo.Tables("BJHeader").Rows(i)("SchDate") = DateAdd(DateInterval.Hour, +8, BJInfo.Tables("BJHeader").Rows(i)("SchDate"))
                                sqlstr = String.Format("Insert into T_PWCBJHeader (BJ, OrgCode, DJ, DJQty, DJReleasedDate, TLA, ProdLine, PWC, BPA, BJQty, OpenQty, Status, ScheduleDate, PWCSubInv, PWCLocator, TLASubInv, TLALocator, CreatedOn, CreatedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', '{17}', '{18}', '{19}')", BJInfo.Tables("BJHeader").Rows(i)("BJ"), OracleLoginData.OrgCode, BJInfo.Tables("BJHeader").Rows(i)("DJ"), BJInfo.Tables("BJHeader").Rows(i)("BJQty"), BJInfo.Tables("BJHeader").Rows(i)("DJReleasedDate"), BJInfo.Tables("BJHeader").Rows(i)("TLA"), BJInfo.Tables("BJHeader").Rows(i)("ProdLine"), BJInfo.Tables("BJHeader").Rows(i)("PWC"), BJInfo.Tables("BJHeader").Rows(i)("BPA"), BJInfo.Tables("BJHeader").Rows(i)("BJQty"), BJInfo.Tables("BJHeader").Rows(i)("BJQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty"), BJInfo.Tables("BJHeader").Rows(i)("Status"), BJInfo.Tables("BJHeader").Rows(i)("SchDate"), BJInfo.Tables("BJHeader").Rows(i)("PWCSubInv"), BJInfo.Tables("BJHeader").Rows(i)("PWCLocator"), BJInfo.Tables("BJHeader").Rows(i)("TLASubInv"), BJInfo.Tables("BJHeader").Rows(i)("TLALocator"), DateTime.Now, BJInfo.Tables("BJHeader").Rows(i)("CreatedBy"), FixNull(BJInfo.Tables("BJHeader").Rows(i)("Remarks")))
                                da.ExecuteNonQuery(sqlstr)
                            Next

                            For i = 0 To BJInfo.Tables("BJItem").Rows.Count - 1
                                sqlstr = String.Format("Insert into T_PWCBJItem (BJ, Component, ProcessCode, Usage, SupplyType, MO, ChangedOn, ChangedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", BJInfo.Tables("BJItem").Rows(i)("BJ"), BJInfo.Tables("BJItem").Rows(i)("Item"), BJInfo.Tables("BJItem").Rows(i)("ProcessCode"), BJInfo.Tables("BJItem").Rows(i)("Usage"), BJInfo.Tables("BJItem").Rows(i)("SupplyType"), BJInfo.Tables("BJItem").Rows(i)("MoveOrder"), DateTime.Now, BJInfo.Tables("BJItem").Rows(i)("ChangedBy"), FixNull(BJInfo.Tables("BJItem").Rows(i)("Remarks")))
                                da.ExecuteNonQuery(sqlstr)
                            Next
                        Else
                            BJ_SaveChange = "BJs already generated for this DJ by somebody."
                            Exit Function
                        End If
                    Else

                    End If
                Else
                    ' Change BJ
                    For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                        BJ = BJInfo.Tables("BJHeader").Rows(i)("BJ")
                        DJ = BJInfo.Tables("BJHeader").Rows(i)("DJ")
                        BPA = BJInfo.Tables("BJHeader").Rows(i)("BPA")
                        OpenQty = BJInfo.Tables("BJHeader").Rows(i)("BJQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty")
                        Status = BJInfo.Tables("BJHeader").Rows(i)("Status").ToString
                        SchDate = DateAdd(DateInterval.Hour, +8, BJInfo.Tables("BJHeader").Rows(i)("SchDate"))
                        Remarks = FixNull(BJInfo.Tables("BJHeader").Rows(i)("Remarks"))
                        ChangedBy = FixNull(BJInfo.Tables("BJHeader").Rows(i)("ChangedBy"))

                        If Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is Nothing AndAlso Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is DBNull.Value Then
                            'ChangedOn = BJInfo.Tables("BJHeader").Rows(i)("ChangedOn")
                            ChangedOn = DateTime.Now
                            sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}', ChangedOn = '{7}', ChangedBy = '{8}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA, ChangedOn, ChangedBy)
                            da.ExecuteNonQuery(sqlstr)
                        Else
                            sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA)
                            da.ExecuteNonQuery(sqlstr)
                        End If
                    Next
                End If
                BJ_SaveChange = ""

                'If txtDJ <> "" Then   ' txtDJ is not blank, insert or update
                '    sqlstr = String.Format("select top (1) BJ from T_PWCBJHeader where OrgCode = '{0}' AND DJ = '{1}'", OracleLoginData.OrgCode, txtDJ)
                '    BJ = FixNull(da.ExecuteScalar(sqlstr))

                '    If BJ = "" Then
                '        'Create BJ
                '        For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                '            BJInfo.Tables("BJHeader").Rows(i)("SchDate") = DateAdd(DateInterval.Hour, +8, BJInfo.Tables("BJHeader").Rows(i)("SchDate"))
                '            sqlstr = String.Format("Insert into T_PWCBJHeader (BJ, OrgCode, DJ, DJQty, TLA, ProdLine, PWC, BPA, BJQty, OpenQty, Status, ScheduleDate, CreatedOn, CreatedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}')", BJInfo.Tables("BJHeader").Rows(i)("BJ"), OracleLoginData.OrgCode, BJInfo.Tables("BJHeader").Rows(i)("DJ"), BJInfo.Tables("BJHeader").Rows(i)("BJQty"), BJInfo.Tables("BJHeader").Rows(i)("TLA"), BJInfo.Tables("BJHeader").Rows(i)("ProdLine"), BJInfo.Tables("BJHeader").Rows(i)("PWC"), BJInfo.Tables("BJHeader").Rows(i)("BPA"), BJInfo.Tables("BJHeader").Rows(i)("BJQty"), BJInfo.Tables("BJHeader").Rows(i)("BJQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty"), BJInfo.Tables("BJHeader").Rows(i)("Status"), BJInfo.Tables("BJHeader").Rows(i)("SchDate"), DateTime.Now, BJInfo.Tables("BJHeader").Rows(i)("CreatedBy"), fixnull(BJInfo.Tables("BJHeader").Rows(i)("Remarks")))
                '            da.ExecuteNonQuery(sqlstr)
                '        Next

                '        For i = 0 To BJInfo.Tables("BJItem").Rows.Count - 1
                '            sqlstr = String.Format("Insert into T_PWCBJItem (BJ, Component, ProcessCode, Usage, SupplyType, MO, ChangedOn, ChangedBy, Remarks) Values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", BJInfo.Tables("BJItem").Rows(i)("BJ"), BJInfo.Tables("BJItem").Rows(i)("Item"), BJInfo.Tables("BJItem").Rows(i)("ProcessCode"), BJInfo.Tables("BJItem").Rows(i)("Usage"), BJInfo.Tables("BJItem").Rows(i)("SupplyType"), BJInfo.Tables("BJItem").Rows(i)("MoveOrder"), DateTime.Now, BJInfo.Tables("BJItem").Rows(i)("ChangedBy"), fixnull(BJInfo.Tables("BJItem").Rows(i)("Remarks")))
                '            da.ExecuteNonQuery(sqlstr)
                '        Next
                '    Else
                '        ' Change BJ
                '        For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                '            BJ = BJInfo.Tables("BJHeader").Rows(i)("BJ")
                '            DJ = BJInfo.Tables("BJHeader").Rows(i)("DJ")
                '            BPA = BJInfo.Tables("BJHeader").Rows(i)("BPA")
                '            OpenQty = BJInfo.Tables("BJHeader").Rows(i)("BJQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty")
                '            Status = BJInfo.Tables("BJHeader").Rows(i)("Status").ToString
                '            SchDate = DateAdd(DateInterval.Hour, +8, BJInfo.Tables("BJHeader").Rows(i)("SchDate"))
                '            Remarks = fixnull(BJInfo.Tables("BJHeader").Rows(i)("Remarks"))
                '            ChangedBy = FixNull(BJInfo.Tables("BJHeader").Rows(i)("ChangedBy"))

                '            If Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is Nothing AndAlso Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is DBNull.Value Then
                '                'ChangedOn = BJInfo.Tables("BJHeader").Rows(i)("ChangedOn")
                '                ChangedOn = DateTime.Now
                '                sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}', ChangedOn = '{7}', ChangedBy = '{8}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA, ChangedOn, ChangedBy)
                '                da.ExecuteNonQuery(sqlstr)
                '            Else
                '                sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA)
                '                da.ExecuteNonQuery(sqlstr)
                '            End If
                '        Next
                '    End If

                '    ' txtDJ is blank, only Update
                'Else
                '    ' Change BJ
                '    For i = 0 To BJInfo.Tables("BJHeader").Rows.Count - 1
                '        BJ = BJInfo.Tables("BJHeader").Rows(i)("BJ")
                '        DJ = BJInfo.Tables("BJHeader").Rows(i)("DJ")
                '        BPA = BJInfo.Tables("BJHeader").Rows(i)("BPA")
                '        OpenQty = BJInfo.Tables("BJHeader").Rows(i)("BJQty") - BJInfo.Tables("BJHeader").Rows(i)("CmpQty")
                '        Status = BJInfo.Tables("BJHeader").Rows(i)("Status").ToString
                '        SchDate = DateAdd(DateInterval.Hour, +8, BJInfo.Tables("BJHeader").Rows(i)("SchDate"))
                '        Remarks = fixnull(BJInfo.Tables("BJHeader").Rows(i)("Remarks"))
                '        ChangedBy = FixNull(BJInfo.Tables("BJHeader").Rows(i)("ChangedBy"))

                '        If Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is Nothing AndAlso Not BJInfo.Tables("BJHeader").Rows(i)("ChangedOn") Is DBNull.Value Then
                '            'ChangedOn = BJInfo.Tables("BJHeader").Rows(i)("ChangedOn")
                '            ChangedOn = DateTime.Now
                '            sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}', ChangedOn = '{7}', ChangedBy = '{8}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA, ChangedOn, ChangedBy)
                '            da.ExecuteNonQuery(sqlstr)
                '        Else
                '            sqlstr = String.Format("Update T_PWCBJHeader SET OpenQty = '{0}', Status = '{1}', ScheduleDate='{2}', Remarks = '{3}' where BJ = '{4}' AND DJ = '{5}' AND BPA = '{6}'", OpenQty, Status, SchDate, Remarks, BJ, DJ, BPA)
                '            da.ExecuteNonQuery(sqlstr)
                '        End If
                '    Next
                'End If
                'BJ_SaveChange = ""
            End Using
        Catch ex As Exception
            ErrorLogging("MaterialMovement-BJ_SaveChange", OracleLoginData.User, ex.Message & ex.Source, "E")
            BJ_SaveChange = ex.Message & ex.Source
        End Try
    End Function

    Public Function CheckBJInfo(ByVal DJ As String, ByVal PWC As String, ByVal BJds As DataSet, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Using da As DataAccess = GetDataAccess()
            Dim BJ, sqlstr, sqlstr1, sqlstr2, sql(1), table(1) As String
            Dim i As Integer
            Dim OpenQty As Decimal
            Dim SchDate As DateTime
            Dim dr As DataRow
            Dim dr_del() As DataRow = Nothing
            Dim BJHeader As New DataTable("BJHeader")
            Dim BJItem As New DataTable("BJItem")

            Try
                CheckBJInfo = New BJ_Rs
                'If Not BJHeader Is Nothing AndAlso BJHeader.Rows.Count > 0 Then
                If FixNull(DJ) <> "" Then
                    sqlstr1 = String.Format("Select top (1) ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, ScheduleDate as SchDate, Remarks, PWC, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where DJ = '{0}' AND OrgCode = '{1}'", DJ, OracleLoginData.OrgCode)
                    sqlstr2 = String.Format("Select top (1) Component as Item, ProcessCode, Usage, SupplyType, MO as MoveOrder, 0 as PickedQty, T_PWCBJItem.BJ, T_PWCBJItem.ChangedOn, T_PWCBJItem.ChangedBy, T_PWCBJItem.Remarks from T_PWCBJItem with (nolock) inner join T_PWCBJHeader with (nolock) on T_PWCBJItem.BJ = T_PWCBJHeader.BJ where DJ = '{0}' AND OrgCode = '{1}'", DJ, OracleLoginData.OrgCode)
                End If
                sql(0) = sqlstr1
                sql(1) = sqlstr2
                table(0) = "BJHeader"
                table(1) = "BJItem"

                CheckBJInfo.BJInfo = da.ExecuteDataSet(sql, table)

                If CheckBJInfo.BJInfo.Tables("BJHeader").Rows.Count > 0 Then
                    BJ = FixNull(da.ExecuteScalar(String.Format("Select top (1) BJ from T_PWCBJHeader where DJ = '{0}' AND OrgCode = '{1}' and (BJQty <> OpenQty or Status <> 'Cancelled')", DJ, OracleLoginData.OrgCode)))
                    If BJ <> "" Then
                        CheckBJInfo.ErrMsg = "BJs already exist"
                        CheckBJInfo.Flag = "90"
                    Else
                        CheckBJInfo.ErrMsg = "No BJs found"
                        CheckBJInfo.Flag = "91"
                    End If
                Else
                    'BJ_GetOpenBJ.BJInfo.Tables("BJHeader").Rows.RemoveAt(0)
                    'BJ_GetOpenBJ.BJInfo.Tables("BJItem").Rows.RemoveAt(0)
                    CheckBJInfo.ErrMsg = "No BJs found"
                    CheckBJInfo.Flag = "91"
                End If

                'Else
                'BJ_GetOpenBJ.BJInfo = Nothing
                'BJ_GetOpenBJ.ErrMsg = "No open BJs found"
                'BJ_GetOpenBJ.Flag = "1"
                'End If
                Return CheckBJInfo
            Catch ex As Exception
                ErrorLogging("BJM-CheckBJInfo", OracleLoginData.User, ex.Message & ex.Source, "E")
                CheckBJInfo.BJInfo = Nothing
                CheckBJInfo.Flag = "92"
                CheckBJInfo.ErrMsg = ex.Message & ex.Source
                Return CheckBJInfo
            End Try
        End Using
    End Function

    Public Function BJ_GetBJ(ByVal DJ As String, ByVal PWC As String, ByVal BJds As DataSet, ByVal OnlyOpen As Boolean, ByVal OracleLoginData As ERPLogin) As BJ_Rs
        Using da As DataAccess = GetDataAccess()
            Dim sqlstr, sqlstr1, sqlstr2, sql(1), table(1) As String
            Dim i As Integer
            Dim OpenQty As Decimal
            Dim SchDate As DateTime
            Dim dr As DataRow
            Dim dr_del() As DataRow = Nothing
            Dim BJHeader As New DataTable("BJHeader")
            Dim BJItem As New DataTable("BJItem")

            Try
                'Get BJHeader

                'If DJ = "" Then
                '    sqlstr = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, ScheduleDate as SchDate, Remarks, PWC, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where PWC = '{0}' AND Status not in ('Completed', 'Cancelled')", PWC)
                'Else
                '    sqlstr1 = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, ScheduleDate as SchDate, Remarks, PWC, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where DJ = '{0}' AND PWC = '{1}' AND Status not in ('Completed', 'Cancelled')", DJ, PWC)
                'End If
                'BJHeader = da.ExecuteDataTable(sqlstr)

                BJ_GetBJ = New BJ_Rs
                'If Not BJHeader Is Nothing AndAlso BJHeader.Rows.Count > 0 Then
                If OnlyOpen = True Then
                    If FixNull(DJ) <> "" Then
                        sqlstr1 = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, DateAdd(Hour, -8, ScheduleDate) as SchDate, Remarks, DJReleasedDate, PWC, PWCSubInv, PWCLocator, TLASubInv, TLALocator, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where DJ = '{0}' AND PWC = '{1}' AND OrgCode = '{2}' AND Status not in ('Completed', 'Cancelled')", DJ, PWC, OracleLoginData.OrgCode)
                        sqlstr2 = String.Format("Select Component as Item, ProcessCode, Usage, (Usage * T_PWCBJHeader.BJQty) as ReqQty, SupplyType, MO as MoveOrder, 0 as PickedQty, T_PWCBJItem.BJ, T_PWCBJItem.ChangedOn, T_PWCBJItem.ChangedBy, T_PWCBJItem.Remarks from T_PWCBJItem with (nolock) inner join T_PWCBJHeader with (nolock) on T_PWCBJItem.BJ = T_PWCBJHeader.BJ where DJ = '{0}' AND PWC = '{1}' AND OrgCode = '{2}' AND Status not in ('Completed', 'Cancelled')", DJ, PWC, OracleLoginData.OrgCode)
                    Else
                        sqlstr1 = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, DateAdd(Hour, -8, ScheduleDate) as SchDate, Remarks, DJReleasedDate, PWC, PWCSubInv, PWCLocator, TLASubInv, TLALocator, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where OrgCode = '{0}' AND PWC = '{1}' AND Status not in ('Completed', 'Cancelled')", OracleLoginData.OrgCode, PWC)
                        sqlstr2 = String.Format("Select Component as Item, ProcessCode, Usage, (Usage * T_PWCBJHeader.BJQty) as ReqQty, SupplyType, MO as MoveOrder, 0 as PickedQty, T_PWCBJItem.BJ, T_PWCBJItem.ChangedOn, T_PWCBJItem.ChangedBy, T_PWCBJItem.Remarks from T_PWCBJItem with (nolock) inner join T_PWCBJHeader with (nolock) on T_PWCBJItem.BJ = T_PWCBJHeader.BJ where OrgCode = '{0}' AND PWC = '{1}' AND Status not in ('Completed', 'Cancelled')", OracleLoginData.OrgCode, PWC)
                    End If
                Else
                    If FixNull(DJ) <> "" Then
                        sqlstr1 = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, DateAdd(Hour, -8, ScheduleDate) as SchDate, Remarks, DJReleasedDate, PWC, PWCSubInv, PWCLocator, TLASubInv, TLALocator, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where DJ = '{0}' AND PWC = '{1}' AND OrgCode = '{2}'", DJ, PWC, OracleLoginData.OrgCode)
                        sqlstr2 = String.Format("Select Component as Item, ProcessCode, Usage, (Usage * T_PWCBJHeader.BJQty) as ReqQty, SupplyType, MO as MoveOrder, 0 as PickedQty, T_PWCBJItem.BJ, T_PWCBJItem.ChangedOn, T_PWCBJItem.ChangedBy, T_PWCBJItem.Remarks from T_PWCBJItem with (nolock) inner join T_PWCBJHeader with (nolock) on T_PWCBJItem.BJ = T_PWCBJHeader.BJ where DJ = '{0}' AND PWC = '{1}' AND OrgCode = '{2}'", DJ, PWC, OracleLoginData.OrgCode)
                    Else
                        sqlstr1 = String.Format("Select ProdLine, DJ, TLA, BJ, BPA, BJQty, (BJQty - OpenQty) as CmpQty, 0 as CurCmpQty, Status, DateAdd(Hour, -8, ScheduleDate) as SchDate, Remarks, DJReleasedDate, PWC, PWCSubInv, PWCLocator, TLASubInv, TLALocator, CreatedOn, CreatedBy, ChangedOn, ChangedBy, OpenQty from T_PWCBJHeader with (nolock) where OrgCode = '{0}' AND PWC = '{1}'", OracleLoginData.OrgCode, PWC)
                        sqlstr2 = String.Format("Select Component as Item, ProcessCode, Usage, (Usage * T_PWCBJHeader.BJQty) as ReqQty, SupplyType, MO as MoveOrder, 0 as PickedQty, T_PWCBJItem.BJ, T_PWCBJItem.ChangedOn, T_PWCBJItem.ChangedBy, T_PWCBJItem.Remarks from T_PWCBJItem with (nolock) inner join T_PWCBJHeader with (nolock) on T_PWCBJItem.BJ = T_PWCBJHeader.BJ where OrgCode = '{0}' AND PWC = '{1}'", OracleLoginData.OrgCode, PWC)
                    End If
                End If

                sql(0) = sqlstr1
                sql(1) = sqlstr2
                table(0) = "BJHeader"
                table(1) = "BJItem"

                BJ_GetBJ.BJInfo = da.ExecuteDataSet(sql, table)

                If BJ_GetBJ.BJInfo.Tables("BJHeader").Rows.Count > 0 Then
                    BJ_GetBJ.ErrMsg = ""
                    BJ_GetBJ.Flag = "0"
                Else
                    'BJ_GetOpenBJ.BJInfo.Tables("BJHeader").Rows.RemoveAt(0)
                    'BJ_GetOpenBJ.BJInfo.Tables("BJItem").Rows.RemoveAt(0)
                    BJ_GetBJ.ErrMsg = "No open BJs found"
                    BJ_GetBJ.Flag = "1"
                End If

                'Else
                'BJ_GetOpenBJ.BJInfo = Nothing
                'BJ_GetOpenBJ.ErrMsg = "No open BJs found"
                'BJ_GetOpenBJ.Flag = "1"
                'End If
                Return BJ_GetBJ
            Catch ex As Exception
                ErrorLogging("MaterialMovement-BJ_SaveChange", OracleLoginData.User, ex.Message & ex.Source, "E")
                BJ_GetBJ.BJInfo = Nothing
                BJ_GetBJ.Flag = "2"
                BJ_GetBJ.ErrMsg = ex.Message & ex.Source
                Return BJ_GetBJ
            End Try
        End Using
    End Function

    Public Function BJ_GenMO(ByVal PWC As String, ByVal DJ As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dsReq As DataSet
            Dim dt As DataTable
            Dim drow() As DataRow = Nothing
            Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()
            Dim TransactionID As Long
            Dim OrgID As String
            Dim i As Integer

            Try
                dsReq = New DataSet
                strSQL = String.Format("ora_get_bj_mo '{0}'", OracleLoginData.OrgCode)
                dsReq = da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("BJ_GenMO-Ora_get_bj_mo", "", ex.Message & ex.Source, "E")
                BJ_GenMO = ex.Message & ex.Source.ToString
                Return BJ_GenMO
            End Try

            For i = 0 To dsReq.Tables(0).Rows.Count - 1
                If dsReq.Tables(0).Rows(i).RowState = DataRowState.Unchanged Then
                    dsReq.Tables(0).Rows(i).SetAdded()
                End If
            Next

            Try
                OrgID = GetOrgID(OracleLoginData.OrgCode)

                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_invcc_pkg.insert_bj_mo_data"
                comm.Parameters.Add("p_org_code", OracleType.VarChar, 240).SourceColumn = "orgcode"
                comm.Parameters.Add("p_item_num", OracleType.VarChar, 240).SourceColumn = "component"
                comm.Parameters.Add("p_locator", OracleType.VarChar, 240).SourceColumn = "pwclocator"
                comm.Parameters.Add("p_qty", OracleType.Double).SourceColumn = "qty"
                comm.Parameters.Add("o_flag", OracleType.VarChar, 240).SourceColumn = "o_flag"
                comm.Parameters.Add("o_msg", OracleType.VarChar, 240).SourceColumn = "o_msg"
                comm.Parameters("o_flag").Direction = ParameterDirection.Output
                comm.Parameters("o_msg").Direction = ParameterDirection.Output

                oda_h.InsertCommand = comm
                oda_h.Update(dsReq.Tables(0))

                Dim result_flag As String
                drow = dsReq.Tables(0).Select("o_flag <> 'Y'")

                If drow.Length = 0 Then
                    result_flag = submit_mo_list("BJ_SaveChange", PWC, OracleLoginData)

                    If result_flag = "Y" Then
                        BJ_GenMO = ""
                        Return BJ_GenMO
                    Else
                        ErrorLogging("BJ_GenMO-Insert_bj_mo_data", OracleLoginData.User, result_flag, "I")
                        BJ_GenMO = result_flag
                        Return BJ_GenMO
                    End If
                Else
                    For i = 0 To (drow.Length - 1)
                        result_flag = result_flag & drow(i)("o_msg").ToString & " "
                        ErrorLogging("BJ_GenMO-Insert_bj_mo_data", OracleLoginData.User, drow(i)("o_msg").ToString, "I")
                    Next
                    BJ_GenMO = result_flag
                    Return BJ_GenMO
                End If
                'Return BJ_GenMO
            Catch ex As Exception
                ErrorLogging("BJ_GenMO-Insert_bj_mo_data", OracleLoginData.User, ex.Message & ex.Source, "E")
                BJ_GenMO = ex.Message & ex.Source.ToString
                Return BJ_GenMO
            End Try
        End Using
    End Function

    Public Function submit_mo_list(ByVal MoveType As String, ByVal PWC As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim a As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()
            Dim aa As OracleString
            Dim resp As Integer = OracleLoginData.RespID_Inv
            Dim appl As Integer = OracleLoginData.AppID_Inv

            Try
                Oda.CommandType = CommandType.StoredProcedure

                Oda.CommandText = "apps.xxetr_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
                Oda.Connection.Open()
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()

                Oda.CommandText = "apps.xxetr_invcc_pkg.create_bjmo_submit_req"
                Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OracleLoginData.OrgCode
                Oda.Parameters.Add("p_pwc", OracleType.VarChar, 240).Value = PWC
                Oda.Parameters.Add("p_processing_mode_code", OracleType.VarChar, 240).Value = "MANUAL"

                Oda.Parameters.Add("o_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_msg", OracleType.VarChar, 240).Direction = ParameterDirection.Output

                'Oda.Connection.Open()
                a = CInt(Oda.ExecuteNonQuery())
                MsgFlag = FixNull(Oda.Parameters("o_flag").Value)
                If MsgFlag = "Y" Then

                Else
                    MsgFlag = FixNull(Oda.Parameters("o_msg").Value)
                End If
                Return DirectCast(MsgFlag, String)
                Oda.Connection.Close()
            Catch oe As OracleException
                ErrorLogging(MoveType, OracleLoginData.User, oe.Message & oe.Source, "E")
                Return oe.Message & oe.Source.ToString
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

#Region "WFOE U-Turn"
    Public Function CheckUTurnSubinv(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim Sqlstr As String

        Try
            Using da As DataAccess = GetDataAccess()
                Sqlstr = String.Format("Select ConfigID,Value,Exempt from T_Config with (nolock) where ConfigID in ('HHF016','HHF017','HHF019')")
                CheckUTurnSubinv = da.ExecuteDataSet(Sqlstr)
            End Using

        Catch ex As Exception
            ErrorLogging("CheckUTurnSubinv", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function ReadCLID_UTurn(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim Sqlstr As String
        Dim ds, dsCLID As New DataSet
        Dim Cnt_ID As Integer

        Try
            Using da As DataAccess = GetDataAccess()

                Sqlstr = String.Format("Select OrgCode, CLID, MaterialNo, MaterialRevision, QtyBaseUOM, BaseUOM, StatusCode, CONVERT(varchar, ExpDate, 101) as Expdate, RecDocNo as RTNo, RTLot, RoHS, SLOC, StorageBin, StorageType, BoxID, 1 as Cnt_CLID  from T_CLMaster with (nolock) where CLID = '{0}'", CLID)
                dsCLID = da.ExecuteDataSet(Sqlstr, "CLIDS")

                If dsCLID Is Nothing OrElse dsCLID.Tables.Count < 1 OrElse dsCLID.Tables(0).Rows.Count < 1 Then
                    ds = New DataSet
                    Sqlstr = String.Format("Select CLID from T_CLMaster with (nolock) where StatusCode = 1 and BoxID = '{0}'", CLID)
                    ds = da.ExecuteDataSet(Sqlstr, "Cnt")
                    If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                        Cnt_ID = ds.Tables(0).Rows.Count
                    Else
                        Cnt_ID = 0
                    End If

                    Sqlstr = String.Format("Select OrgCode, CLID, MaterialNo, MaterialRevision, QtyBaseUOM, BaseUOM, StatusCode, CONVERT(varchar, ExpDate, 101) as Expdate, RecDocNo as RTNo, RTLot, RoHS, SLOC, StorageBin, StorageType, BoxID, '{1}' as Cnt_CLID  from T_CLMaster with (nolock) where BoxID = '{0}'", CLID, Cnt_ID)
                    dsCLID = da.ExecuteDataSet(Sqlstr, "CLIDS")
                End If
                Return dsCLID
            End Using

        Catch ex As Exception
            ErrorLogging("CheckUTurnSubinv", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function UTurnDelivery(ByVal OracleLoginData As ERPLogin, ByVal CLIDS As DataSet) As String
        Dim Sqlstr As String

        Try
            Using da As DataAccess = GetDataAccess()
                Sqlstr = String.Format("exec sp_UTurn_Delivery N'{0}', '{1}',N'{2}'", DStoXML(CLIDS), OracleLoginData.OrgCode, OracleLoginData.User)
                da.ExecuteNonQuery(Sqlstr)

                UTurnDelivery = ""
            End Using

        Catch ex As Exception
            ErrorLogging("UTurnDelivery", OracleLoginData.User, ex.Message & ex.Source, "E")
            UTurnDelivery = ex.Message & ex.Source
        End Try
    End Function

    Public Function CheckUTurnCLIDFormat(ByVal ExcelData As DataSet, ByVal SubInv As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim i As Integer
        Dim DR_Blank(), dr_check() As DataRow
        Dim strSQL, SplitSubinv(), Flag As String
        Dim dsResult As DataSet

        Try
            'deal with blank rows
            DR_Blank = Nothing
            DR_Blank = ExcelData.Tables(0).Select("(CLID is null or clid = '')")
            If DR_Blank.Length > 0 Then
                For Each drblank As DataRow In DR_Blank
                    ExcelData.Tables(0).Rows.Remove(drblank)
                Next
            End If

            dsResult = New DataSet
            Using da As DataAccess = GetDataAccess()

                Try
                    strSQL = String.Format("exec sp_CheckUTurnCLIDFormat N'{0}', '{1}','{2}'", DStoXML(ExcelData), SubInv, OracleERPLogin.User)
                    dsResult = da.ExecuteDataSet(strSQL)
                Catch ex As Exception
                    ErrorLogging("MateralMovement-UploadEMCfile", "", ex.Message & ex.Source, "E")
                End Try
            End Using

            Erase SplitSubinv
            SplitSubinv = Split(SubInv, ",")

            For Each dr As DataRow In ExcelData.Tables(0).Rows
                dr_check = Nothing
                dr_check = dsResult.Tables(0).Select("CLID = '" & dr("CLID") & "'")

                Flag = "N"
                If dr_check.Length > 0 Then
                    If dr_check(0)("OrgCode") <> OracleERPLogin.OrgCode Then
                        dr_check(0)("ErrorMsg") = "Error: incorrect OrgCode"
                    End If

                    If dr_check(0)("StatusCode") <> 6 Then
                        dr_check(0)("ErrorMsg") = FixNull(dr_check(0)("ErrorMsg")) & "Error: incorrect CLID StatusCode!"
                    End If

                    'Check if subinventory match
                    If InStr(SubInv, ",") > 0 Then
                        If Not InStr(Join(SplitSubinv, ","), FixNull(dr_check(0)("Subinventory"))) > 0 Then
                            'If Not InStr(SubInv, dr_check(0)("Subinventory")) > 0 Then
                            dr_check(0)("ErrorMsg") = FixNull(dr_check(0)("ErrorMsg")) & "Error: Subinventory does not match!"
                        End If
                    Else
                        If SubInv <> FixNull(dr_check(0)("Subinventory")) Then
                            'If Not InStr(SubInv, dr_check(0)("Subinventory")) > 0 Then
                            dr_check(0)("ErrorMsg") = FixNull(dr_check(0)("ErrorMsg")) & "Error: Subinventory does not match!"
                        End If
                    End If


                    'For i = 0 To SplitSubinv.Length - 1
                    '    If dr_check(0)("Subinventory") = SplitSubinv(i) Then
                    '        Flag = "Y"
                    '    End If
                    'Next
                    'If Flag = "N" Then
                    '    dr_check(0)("ErrMsg") = "Error: Subinventory does not match!"
                    'End If

                Else
                    myDataRow = dsResult.Tables(0).NewRow()
                    myDataRow("CLID") = dr("CLID")
                    myDataRow("ErrorMsg") = "Error: CLID does Not exist!"
                    dsResult.Tables(0).Rows.Add(myDataRow)
                End If
            Next

            Return dsResult

        Catch ex As Exception
            ErrorLogging("UTurnDelivery-CheckUTurnCLIDFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function PostUTurnStatusChange(ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As String
        Dim RowIndex As Integer = 0
        Dim SQL As String
        Dim Result As String

        Try
            Using da As DataAccess = GetDataAccess()

                Try
                    SQL = String.Format("exec sp_PostUTurnStatusChange N'{0}', '{1}'", DStoXML(ExcelData), OracleERPLogin.User)
                    Result = Convert.ToString(da.ExecuteScalar(SQL))
                Catch ex As Exception
                    ErrorLogging("UTurnDelivery-PostUTurnStatusChange", "", ex.Message & ex.Source, "E")
                End Try
            End Using

            Return Result

        Catch ex As Exception
            ErrorLogging("UTurnDelivery-PostUTurnStatusChange", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function CheckLocRTList(ByVal Item As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim Sqlstr As String

        Try
            Using da As DataAccess = GetDataAccess()
                Sqlstr = String.Format("select distinct RTLot from t_clmaster with (nolock) where OrgCode = '{0}' and StatusCode in ('1','6') and MaterialNo = '{1}' and SLOC = '{2}' and StorageBin = '{3}'", OracleLoginData.OrgCode, Item, Subinv, Locator)
                CheckLocRTList = da.ExecuteDataSet(Sqlstr)
            End Using

        Catch ex As Exception
            ErrorLogging("CheckLocRTList", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function


#End Region

#Region "Batch Disable CLIDs"
    Public Function CheckCLIDDisableFlag(ByVal OracleLoginData As ERPLogin) As String
        Dim Sqlstr As String
        Dim CLIDConfig As String = ""

        Try
            Using da As DataAccess = GetDataAccess()
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID001'")
                CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr))

                'If Disable / Enable CLID function is not allowed in T_Config, then check user account in table field Exempt
                If CLIDConfig = "NO" Then
                    Sqlstr = String.Format("Select Exempt from T_Config with (nolock) where ConfigID = 'CLID001'")
                    Dim UserExempt As String = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                    If UserExempt <> "" Then
                        If UserExempt.Contains(OracleLoginData.User.ToUpper) Then CLIDConfig = "YES"
                    End If
                End If
                CheckCLIDDisableFlag = CLIDConfig
            End Using

        Catch ex As Exception
            ErrorLogging("CheckCLIDDisableFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function ValidateSubinv(ByVal OracleLoginData As ERPLogin, ByVal SubInv As String) As String
        Dim SubinvList() As String
        Dim i As Integer
        Dim Rst As String
        Try
            ValidateSubinv = ""
            If InStr(SubInv, ",") > 0 Then
                SubinvList = Split(SubInv, ",")

                For i = 0 To SubinvList.Length - 1
                    Rst = ""
                    Rst = ValidateSubLoc(OracleLoginData, SubinvList(i), "")
                    If Rst <> "Y" AndAlso Rst <> "N" Then
                        If ValidateSubinv = "" Then
                            ValidateSubinv = SubinvList(i) & ": " & "^BMR-78@"
                        Else
                            ValidateSubinv = ValidateSubinv & " " & SubinvList(i) & ": " & "^BMR-78@"
                        End If
                    End If
                Next
            Else
                Rst = ""
                Rst = ValidateSubLoc(OracleLoginData, SubInv, "")
                If Rst <> "Y" AndAlso Rst <> "N" Then
                    ValidateSubinv = SubInv & ": " & "^BMR-78@"
                End If
            End If
        Catch ex As Exception
            ErrorLogging("ValidateSubinv", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function CheckBatchDisableCLIDFormat(ByVal ExcelData As DataSet, ByVal SubInv As String, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim i As Integer
        Dim DR_Blank(), dr_check() As DataRow
        Dim strSQL, SplitSubinv(), Flag As String
        Dim dsResult As DataSet

        Try
            'deal with blank rows
            DR_Blank = Nothing
            DR_Blank = ExcelData.Tables(0).Select("(CLID is null or clid = '')")
            If DR_Blank.Length > 0 Then
                For Each drblank As DataRow In DR_Blank
                    ExcelData.Tables(0).Rows.Remove(drblank)
                Next
            End If

            dsResult = New DataSet
            Using da As DataAccess = GetDataAccess()

                Try
                    strSQL = String.Format("exec sp_CheckBatchDisableCLIDFormat N'{0}', '{1}','{2}'", DStoXML(ExcelData), SubInv, OracleERPLogin.User)
                    dsResult = da.ExecuteDataSet(strSQL)
                Catch ex As Exception
                    ErrorLogging("MateralMovement-UploadEMCfile", "", ex.Message & ex.Source, "E")
                End Try
            End Using

            Erase SplitSubinv
            SplitSubinv = Split(SubInv, ",")

            For Each dr As DataRow In ExcelData.Tables(0).Rows
                dr_check = Nothing
                dr_check = dsResult.Tables(0).Select("CLID = '" & dr("CLID") & "'")

                Flag = "N"
                If dr_check.Length > 0 Then
                    If dr_check(0)("StatusCode") <> 1 Then
                        dr_check(0)("ErrorMsg") = "Error: incorrect CLID StatusCode!"
                    End If

                    'Check if subinventory match
                    If InStr(SubInv, ",") > 0 Then
                        If Not InStr(Join(SplitSubinv, ","), FixNull(dr_check(0)("Subinventory"))) > 0 Then
                            'If Not InStr(SubInv, dr_check(0)("Subinventory")) > 0 Then
                            dr_check(0)("ErrorMsg") = FixNull(dr_check(0)("ErrorMsg")) & "Error: Subinventory does not match!"
                        End If
                    Else
                        If SubInv <> FixNull(dr_check(0)("Subinventory")) Then
                            'If Not InStr(SubInv, dr_check(0)("Subinventory")) > 0 Then
                            dr_check(0)("ErrorMsg") = FixNull(dr_check(0)("ErrorMsg")) & "Error: Subinventory does not match!"
                        End If
                    End If


                    'For i = 0 To SplitSubinv.Length - 1
                    '    If dr_check(0)("Subinventory") = SplitSubinv(i) Then
                    '        Flag = "Y"
                    '    End If
                    'Next
                    'If Flag = "N" Then
                    '    dr_check(0)("ErrMsg") = "Error: Subinventory does not match!"
                    'End If

                Else
                    myDataRow = dsResult.Tables(0).NewRow()
                    myDataRow("CLID") = dr("CLID")
                    myDataRow("ErrorMsg") = "Error: CLID does Not exist!"
                    dsResult.Tables(0).Rows.Add(myDataRow)
                End If
            Next

            Return dsResult

        Catch ex As Exception
            ErrorLogging("BatchDisableCLID-CheckBatchDisableCLIDFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function PostBatchDisableCLID(ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As String
        Dim RowIndex As Integer = 0
        Dim SQL As String
        Dim Result As String

        Try
            Using da As DataAccess = GetDataAccess()

                Try
                    SQL = String.Format("exec sp_PostBatchDisableCLID N'{0}', '{1}'", DStoXML(ExcelData), OracleERPLogin.User)
                    Result = Convert.ToString(da.ExecuteScalar(SQL))
                Catch ex As Exception
                    ErrorLogging("MateralMovement-UploadEMCfile", "", ex.Message & ex.Source, "E")
                End Try
            End Using

            Return Result

        Catch ex As Exception
            ErrorLogging("BatchDisableCLID-PostBatchDisableCLID", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

#End Region


#Region "UploadEMCfile"

    Public Function UploadEMCfile(ByVal DS As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Try
                strSQL = String.Format("exec sp_EMCfileUpload N'{0}'", DStoXML(DS))
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("MateralMovement-UploadEMCfile", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

#End Region

    'Public Function UpdateCLMaster(ByVal OracleLoginData As ERPLogin) As String
    '    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
    '    Dim SQLTransaction As SqlClient.SqlTransaction
    '    Dim CLMasterSQLCommand As SqlClient.SqlCommand
    '    Dim i, ra As Integer
    '    Dim strCMD1, strCMD2 As String
    '    Dim ds, CLID_ds, SLOC_ds, PN_ds As DataSet
    '    Dim tmp_CLID As String
    '    Using da As DataAccess = GetDataAccess()
    '        Try
    '            ds = New DataSet
    '            strCMD1 = String.Format("Select SLOC, SType, Bin from T_SAPSLOC where MigrationType = 'SAP'")
    '            ds = da.ExecuteDataSet(strCMD1, "SAPType")

    '            myConn.Open()

    '            SQLTransaction = myConn.BeginTransaction
    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                strCMD2 = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = 0, LastTransaction = 'SAPIM_UpdateCLMaster' where StatusCode = 1 and SLOC = '{1}' and StorageType = '{2}' and StorageBin = '{3}' ", OracleLoginData.User, ds.Tables(0).Rows(i)("SLOC"), ds.Tables(0).Rows(i)("SType"), ds.Tables(0).Rows(i)("Bin"))
    '                CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD2, myConn)
    '                CLMasterSQLCommand.Connection = myConn
    '                CLMasterSQLCommand.Transaction = SQLTransaction
    '                CLMasterSQLCommand.CommandTimeout = 1800000
    '                ra = CLMasterSQLCommand.ExecuteNonQuery()
    '            Next

    '            SQLTransaction.Commit()
    '        Catch ex As Exception
    '            SQLTransaction.Rollback()
    '            ErrorLogging("SAPMigration-UpdateCLMaster", "", ex.Message & ex.Source, "E")
    '            UpdateCLMaster = "Step 1: Failed to set StatusCode to be ZERO for CLID with Migration Type as 'SAP' "
    '            Exit Function
    '        Finally
    '            If myConn.State <> ConnectionState.Closed Then myConn.Close()
    '        End Try

    '        Try
    '            ds = New DataSet
    '            strCMD1 = String.Format("Select CLID from T_CLMASTER where StatusCode = 1 and MigrationStatus is NULL")
    '            ds = da.ExecuteDataSet(strCMD1, "CLIDList")

    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                Try
    '                    CLID_ds = New DataSet
    '                    tmp_CLID = ds.Tables(0).Rows(i)(0).ToString
    '                    strCMD1 = String.Format("Select MaterialNo,QtyBaseUOM,BaseUOM,SLOC,StorageType,StorageBin from T_CLMASTER where CLID = tmp_CLID")
    '                    CLID_ds = da.ExecuteDataSet(strCMD1, "CLIDList")

    '                    SLOC_ds = New DataSet
    '                    strCMD1 = String.Format("Select Subinv,Locator from T_SAPSLOC where SLOC = '{0}' AND SType = '{1}' and Bin = '{2}'", CLID_ds.Tables(0).Rows(0)("SLOC"), CLID_ds.Tables(0).Rows(0)("StorageType"), CLID_ds.Tables(0).Rows(0)("StorageBin"))
    '                    SLOC_ds = da.ExecuteDataSet(strCMD1, "SLOC")

    '                    PN_ds = New DataSet
    '                    strCMD1 = String.Format("Select * from T_SAPPN where SAPPN = '{0}'", CLID_ds.Tables(0).Rows(0)("MaterialNo"))
    '                    PN_ds = da.ExecuteDataSet(strCMD1, "PN")

    '                    strCMD1 = String.Format("Update T_CLMaster set MaterialNo='{0}',MaterialDesc='{1}', QtyBaseUOM='{2}', BaseUOM='{3}',ChangedOn=getdate(),ChangedBy='{4}' where CLID='{1}'")
    '                    da.ExecuteNonQuery(strCMD1)

    '                Catch ex As Exception

    '                End Try
    '            Next

    '        Catch ex As Exception


    '        Finally
    '            If myConn.State <> ConnectionState.Closed Then myConn.Close()
    '        End Try

    '    End Using

    'End Function

    'Public Function GetMaterialInfoByCLID(ByVal CLID As String) As DataSet
    '    Dim ds As New DataSet
    '    Dim Sqlstr1 As String
    '    Using da As DataAccess = GetDataAccess()
    '        Try
    '            Sqlstr1 = String.Format("exec sp_getMaterialInfoByCLID '{0}'", CLID)
    '            ds = da.ExecuteDataSet(Sqlstr1)
    '        Catch ex As Exception
    '            ErrorLogging("GetMaterialInfoByCLID", "", ex.Message & ex.Source, "E")
    '        End Try
    '        Return ds
    '    End Using
    'End Function


    'Public Function GetMaterialToXMLByCLID(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal Pn As String) As String

    '    GetMaterialToXMLByCLID = "Fail to create xml file"
    '    Dim xmlfilename As String
    '    Dim ds As New DataSet()
    '    Dim Sqlstr1 As String
    '    Dim i As Integer
    '    Dim dr As DataRow
    '    xmlfilename = Guid.NewGuid.ToString + ".xml"
    '    Using da As DataAccess = GetDataAccess()
    '        Try
    '            'Sqlstr1 = String.Format("select MaterialNo AS PartNumber, CLID as AGI, LotNo AS LotNumber, QtyBaseUOM AS Quantity, (StatusCode + 1) % 2 AS StatusCode from T_CLmaster where createdon between '{0}' and '{1}' ", From_MR, To_MR)
    '            Sqlstr1 = String.Format("exec sp_getMaterialInfosBydate '{0}' ,'{1}','{2}'", From_MR, To_MR, Pn)
    '            ds = da.ExecuteDataSet(Sqlstr1, "InventoryItemData")
    '            'ds.DataSetName = "InventoryData"
    '            If (ds IsNot Nothing And ds.Tables("InventoryItemData").Rows.Count > 0) Then



    '                Dim myTW As New XmlTextWriter(ConfigurationSettings.AppSettings.Item("XMLFolder") + xmlfilename, Nothing)

    '                myTW.WriteStartDocument()
    '                'myTW.Formatting = Formatting.Indented


    '                myTW.WriteStartElement("InventoryData")
    '                For i = 0 To ds.Tables("InventoryItemData").Rows.Count - 1
    '                    dr = ds.Tables("InventoryItemData").Rows(i)
    '                    myTW.WriteStartElement("InventoryItemData")
    '                    myTW.WriteAttributeString("PartNumber", dr(0).ToString)
    '                    myTW.WriteAttributeString("AGI", dr(1).ToString)
    '                    myTW.WriteAttributeString("LotNumber", dr(2).ToString)
    '                    myTW.WriteAttributeString("Quantity", dr(3).ToString)
    '                    myTW.WriteAttributeString("StatusCode", dr(4).ToString)
    '                    myTW.WriteFullEndElement()
    '                Next
    '                myTW.WriteFullEndElement()
    '                myTW.WriteEndDocument()
    '                myTW.Flush()
    '                myTW.Close()


    '                'ds.WriteXml(ConfigurationSettings.AppSettings.Item("XMLFolder") + xmlfilename)
    '                GetMaterialToXMLByCLID = xmlfilename
    '            End If
    '        Catch ex As Exception
    '            GetMaterialToXMLByCLID = GetMaterialToXMLByCLID + " " + ex.Message
    '            ErrorLogging("GetMaterialToXMLByCLID", "", ex.Message & ex.Source, "E")
    '        End Try

    '    End Using
    'End Function
End Class
