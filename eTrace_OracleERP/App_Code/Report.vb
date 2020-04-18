Imports Microsoft.VisualBasic
Imports system.Data
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.xml

'Added by james hu on Aug 20, 2010
Public Structure ComparisonSign
    Dim More As Boolean     '>=
    Dim eQual As Boolean    '=
    Dim Less As Boolean     '<=
End Structure

Public Class Report
    Inherits PublicFunction

    Public Function ValidReqlineStatus(ByVal p_orgcode As String, ByVal p_dnpo As String, ByVal p_input_type As String, ByVal p_ejit_id As Integer) As String
        Using da As DataAccess = GetDataAccess()
            Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()
            Dim status As String

            Try
                If comm.Connection.State = ConnectionState.Closed Then
                    comm.Connection.Open()
                End If
                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_kanban_pkg.valid_reqline_status"
                comm.Parameters.Add("o_status", OracleType.VarChar, 30).Direction = ParameterDirection.Output
                comm.Parameters.Add("p_org_code", OracleType.VarChar, 10).Value = p_orgcode
                comm.Parameters.Add("p_dnpo", OracleType.VarChar, 30).Value = p_dnpo
                comm.Parameters.Add("p_input_type", OracleType.VarChar, 30).Value = p_input_type
                comm.Parameters.Add("p_ejit_id", OracleType.Int32).Value = p_ejit_id

                comm.ExecuteNonQuery()
                status = comm.Parameters("o_status").Value.ToString()

            Catch ex As Exception
                ErrorLogging("ValidReqlineStatus", "", ex.Message & ex.Source, "E")
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
            Return status
        End Using
    End Function

    Public Function TCLIDData(ByVal CLID As String) As DataSet
        TCLIDData = New DataSet
        Dim GetCLIDTab As DataTable
        Dim myDataColumn As DataColumn

        GetCLIDTab = New Data.DataTable("CLIDTab")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialDesc", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QTYBaseUOM", System.Type.GetType("System.Decimal"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BaseUOM", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)

        TCLIDData.Tables.Add(GetCLIDTab)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select CLID,OrgCode,MaterialNO,MaterialDesc,QtyBaseUOM,BaseUOM,Manufacturer,ManufacturerPN from T_CLMaster with (nolock) where CLID = @CLID", myConn)
        TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        TDHeaderSQLCommand.Parameters("@CLID").Value = CLID

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = GetCLIDTab.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNO") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MaterialDesc") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("QTYBaseUOM") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("BaseUOM") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then GetOneRow("Manufacturer") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then GetOneRow("ManufacturerPN") = objReader.GetValue(7)

                GetCLIDTab.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TCLIDData", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function TCLIDInfo(ByVal CLID As String) As String
        Dim sInfo As String = ""
        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select MaterialNO from T_CLMaster with (nolock) where CLID = @CLID", myConn)
        TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        TDHeaderSQLCommand.Parameters("@CLID").Value = CLID

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()
                sInfo = objReader.GetValue(0)
            End While
            myConn.Close()
        Catch ex As Exception
            sInfo = "Error"
            ErrorLogging("Report-TCLIDData", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return sInfo
    End Function

    Public Function TItemData(ByVal Item As String) As DataSet
        TItemData = New DataSet
        Dim GetItemTab As DataTable
        Dim myDataColumn As DataColumn

        GetItemTab = New Data.DataTable("ItemTab")
        myDataColumn = New Data.DataColumn("ENPCItemNo", System.Type.GetType("System.String"))
        GetItemTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("EPItemNo", System.Type.GetType("System.String"))
        GetItemTab.Columns.Add(myDataColumn)

        TItemData.Tables.Add(GetItemTab)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select ENPCItemNo,EPItemNo from Tmp_Material with (nolock) where ENPCItemNo = @Item", myConn)
        TDHeaderSQLCommand.Parameters.Add("@Item", SqlDbType.VarChar, 20, "Item")
        TDHeaderSQLCommand.Parameters("@Item").Value = Item

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = GetItemTab.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("ENPCItemNo") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("EPItemNo") = objReader.GetValue(1)
                GetItemTab.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TItemData", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function TAging(ByVal Org As String, ByVal SubINV As String, ByVal Item As String, ByVal Comparison As ComparisonSign, Optional ByVal AgingDays As Integer = 0) As DataSet
        TAging = New DataSet
        Dim Aging As DataTable
        Dim myDataColumn As DataColumn
        Dim sSign As String
        If Comparison.More = True Then
            sSign = ">="
        ElseIf Comparison.eQual = True Then
            sSign = "="
        ElseIf Comparison.Less = True Then
            sSign = "<="
        Else
            sSign = "="
        End If
        Aging = New Data.DataTable("Aging")
        myDataColumn = New Data.DataColumn("Org", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Item", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubINV", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Decimal"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("DJCompletionDate", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("AgingDays", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SalesOrderNo", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SalesOrderLine", System.Type.GetType("System.String"))
        Aging.Columns.Add(myDataColumn)

        TAging.Tables.Add(Aging)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim sSql As String
        sSql = "Select Org,Item,SubINV,Locator,CLID,QtyBaseUOM,DJCompletionDate,AgingDays,SalesOrderNo,SalesOrderLine from V_AssemblyAging where 1=1 "
        If Trim(Org) <> "" Then
            sSql = sSql + " and Org=@Org"
        End If
        If Trim(SubINV) <> "" Then
            sSql = sSql + " and SubINV=@SubINV"
        End If
        If Trim(Item) <> "" Then
            sSql = sSql + " and Item=@Item"
        End If
        If Trim(AgingDays.ToString) <> "" Then
            sSql = sSql + " and AgingDays" + sSign + "@AgingDays"
        End If
        TDHeaderSQLCommand = New SqlClient.SqlCommand(sSql, myConn)
        If Trim(Org) <> "" Then
            TDHeaderSQLCommand.Parameters.Add("@Org", SqlDbType.VarChar, 20, "Org")
            TDHeaderSQLCommand.Parameters("@Org").Value = Org
        End If
        If Trim(SubINV) <> "" Then
            TDHeaderSQLCommand.Parameters.Add("@SubINV", SqlDbType.VarChar, 20, "SubINV")
            TDHeaderSQLCommand.Parameters("@SubINV").Value = SubINV
        End If
        If Trim(Item) <> "" Then
            TDHeaderSQLCommand.Parameters.Add("@Item", SqlDbType.VarChar, 20, "Item")
            TDHeaderSQLCommand.Parameters("@Item").Value = Item
        End If
        If Trim(AgingDays.ToString) <> "" Then
            TDHeaderSQLCommand.Parameters.Add("@AgingDays", SqlDbType.Int, 20, "AgingDays")
            TDHeaderSQLCommand.Parameters("@AgingDays").Value = AgingDays
        End If

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = Aging.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("Org") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("Item") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("SubINV") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("Locator") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("QtyBaseUOM") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then GetOneRow("DJCompletionDate") = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then GetOneRow("AgingDays") = objReader.GetValue(7)
                If Not objReader.GetValue(8) Is DBNull.Value Then GetOneRow("SalesOrderNo") = objReader.GetValue(8)
                If Not objReader.GetValue(9) Is DBNull.Value Then GetOneRow("SalesOrderLine") = objReader.GetValue(9)

                Aging.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TAging", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function TDJData(ByVal Org As String, ByVal CLID As String) As DataSet
        TDJData = New DataSet
        Dim DJData As DataTable
        Dim myDataColumn As DataColumn
        DJData = New Data.DataTable("DJData")

        If CLID.Trim = "" Then
            myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("PO", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
        Else
            myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MaterialDesc", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("PO", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("CLIDQty", System.Type.GetType("System.String"))
            DJData.Columns.Add(myDataColumn)
        End If

        TDJData.Tables.Add(DJData)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim sSql As String
        If CLID.Trim = "" Then
            'If Org.Trim = "" Then
            '    sSql = "select distinct a.OrgCode,b.PO from T_CLMaster a,T_PO_CLID b where a.CLID=b.CLID order by a.OrgCode,b.PO"
            'Else
            '    sSql = "select distinct a.OrgCode,b.PO from T_CLMaster a,T_PO_CLID b where a.CLID=b.CLID and a.OrgCode=@OrgCode order by a.OrgCode,b.PO"
            'End If
            sSql = "select distinct PO,Model from T_DJInfo where Status='Released' order by PO"
        Else
            sSql = "select a.OrgCode,a.CLID,a.MaterialNo,a.MaterialDesc,b.PO,b.CLIDQty from T_CLMaster a with (nolock),T_PO_CLID b with (nolock) where a.CLID=b.CLID and b.CLID=@CLID"
        End If

        TDHeaderSQLCommand = New SqlClient.SqlCommand(sSql, myConn)
        If Trim(CLID) <> "" Then
            TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 30, "CLID")
            TDHeaderSQLCommand.Parameters("@CLID").Value = CLID
        Else
            If Trim(Org) <> "" Then
                TDHeaderSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 30, "OrgCode")
                TDHeaderSQLCommand.Parameters("@OrgCode").Value = Org
            End If
        End If
        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = DJData.NewRow()
                If Trim(CLID) <> "" Then
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNo") = objReader.GetValue(2)
                    If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MaterialDesc") = objReader.GetValue(3)
                    If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("PO") = objReader.GetValue(4)
                    If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("CLIDQty") = objReader.GetValue(5)
                Else
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("PO") = objReader.GetValue(1)
                End If

                DJData.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TDJData", "", ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function TCLIDIssueData(ByVal CLID As String, ByVal DJ As String) As DataSet
        TCLIDIssueData = New DataSet
        Dim DJData As DataTable
        Dim myDataColumn As DataColumn
        DJData = New Data.DataTable("DJData")


        myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialDesc", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PO", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CLIDQty", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("DateCode", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LotNo", System.Type.GetType("System.String"))
        DJData.Columns.Add(myDataColumn)

        TCLIDIssueData.Tables.Add(DJData)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim sSql As String


        sSql = "sp_121CLIDIssueData"
        TDHeaderSQLCommand = New SqlClient.SqlCommand(sSql, myConn)
        TDHeaderSQLCommand.CommandType = CommandType.StoredProcedure
        TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 30, "CLID")
        TDHeaderSQLCommand.Parameters.Add("@DJ", SqlDbType.VarChar, 30, "DJ")
        TDHeaderSQLCommand.Parameters("@CLID").Value = CLID
        TDHeaderSQLCommand.Parameters("@DJ").Value = DJ

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = DJData.NewRow()
                If Trim(CLID) <> "" Then
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(1)
                    If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNo") = objReader.GetValue(2)
                    If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MaterialDesc") = objReader.GetValue(3)
                    If Not objReader.GetValue(4) Is DBNull.Value Then GetOneRow("PO") = objReader.GetValue(4)
                    If Not objReader.GetValue(5) Is DBNull.Value Then GetOneRow("CLIDQty") = objReader.GetValue(5)
                    If Not objReader.GetValue(6) Is DBNull.Value Then GetOneRow("DateCode") = objReader.GetValue(6)
                    If Not objReader.GetValue(7) Is DBNull.Value Then GetOneRow("LotNo") = objReader.GetValue(7)
                Else
                    If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(0)
                    If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("PO") = objReader.GetValue(1)
                End If

                DJData.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TCLIDIssueData", "", ex.Message & ex.Source, "E")
        End Try

    End Function


    Public Function TCLIDMSLData(ByVal CLID As String) As DataSet
        TCLIDMSLData = New DataSet
        Dim GetCLIDTab As DataTable
        Dim myDataColumn As DataColumn

        GetCLIDTab = New Data.DataTable("CLIDTab")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MSL", System.Type.GetType("System.String"))
        GetCLIDTab.Columns.Add(myDataColumn)

        TCLIDMSLData.Tables.Add(GetCLIDTab)
        Dim GetOneRow As Data.DataRow

        Dim TDHeaderSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader

        TDHeaderSQLCommand = New SqlClient.SqlCommand("Select CLID,OrgCode,MaterialNO,MSL from T_CLMaster with (nolock) where CLID = @CLID", myConn)
        TDHeaderSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        TDHeaderSQLCommand.Parameters("@CLID").Value = CLID

        Try
            myConn.Open()
            TDHeaderSQLCommand.CommandTimeout = TimeOut_M60
            objReader = TDHeaderSQLCommand.ExecuteReader()
            While objReader.Read()

                GetOneRow = GetCLIDTab.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("OrgCode") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("MaterialNO") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then GetOneRow("MSL") = objReader.GetValue(3)

                GetCLIDTab.Rows.Add(GetOneRow)
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Report-TCLIDMSLData", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function TCLIDReturn(ByVal DJ As String, ByVal CLID As String) As DataSet
        TCLIDReturn = New DataSet
        Dim dtCLIDReturn As DataTable
        Dim myDataColumn As DataColumn

        dtCLIDReturn = New Data.DataTable("CLIDReturn")
        myDataColumn = New Data.DataColumn("PO", System.Type.GetType("System.String"))
        dtCLIDReturn.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        dtCLIDReturn.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReturnStatus", System.Type.GetType("System.String"))
        dtCLIDReturn.Columns.Add(myDataColumn)
        TCLIDReturn.Tables.Add(dtCLIDReturn)

        Dim sSQL As String
        Dim mytransCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim GetOneRow As Data.DataRow

        sSQL = String.Format("select PO,CLID,case isnull(returndate,'') when '' then 'N' else 'Y' end as ReturnStatus from T_PO_CLID with (nolock) where PO='{0}' and CLID='{1}' ", Trim(DJ), Trim(CLID))
        mytransCommand = New SqlClient.SqlCommand(sSQL, myConn)

        Try
            myConn.Open()
            mytransCommand.CommandTimeout = TimeOut_M60
            objReader = mytransCommand.ExecuteReader()
            While objReader.Read()
                GetOneRow = dtCLIDReturn.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then GetOneRow("PO") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then GetOneRow("CLID") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then GetOneRow("ReturnStatus") = objReader.GetValue(2)
                dtCLIDReturn.Rows.Add(GetOneRow)

            End While
            myConn.Close()

        Catch ex As Exception
            ErrorLogging("Report-CLIDReturn", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try

    End Function

    Public Function MPNOnHand(ByVal orgID As String) As DataSet
        Dim strSQL = String.Format("exec sp_GetAllMPNOnHand '{0}'", orgID)
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim DA As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(strSQL, myConn)
        Dim DS As DataSet = New DataSet("DS")
        Try
            DA.SelectCommand.CommandTimeout = TimeOut_M60
            DA.Fill(DS)
        Catch ex As Exception
            ErrorLogging("Report-MPNOnHand", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
        Return DS
    End Function

    Public Function OnHandMFGMPN(ByVal Org As String, Optional ByVal Material As String = "", Optional ByVal SubInv As String = "") As DataSet
        Dim strSQL As String
        Dim sMaterial As String = ""
        While Material.IndexOf(",") > 0
            sMaterial = sMaterial + "'" + Material.Substring(0, Material.IndexOf(",")) + "',"
            Material = Material.Substring(Material.IndexOf(",") + 1)
        End While
        If Material <> "" Then
            sMaterial = sMaterial + "'" + Material.ToString.Trim + "'"
        End If
        If sMaterial = "" And SubInv = "" Then
            strSQL = "select OrgCode,MaterialNo,SubInventory,Locator,LotNo,MFG,MPN,Total from v_onhandmfgmpn where OrgCode='" + Org + "'"
        ElseIf sMaterial = "" And SubInv <> "" Then
            strSQL = "select OrgCode,MaterialNo,SubInventory,Locator,LotNo,MFG,MPN,Total from v_onhandmfgmpn where OrgCode='" + Org + "' and SubInv='" + SubInv + "'"
        ElseIf sMaterial <> "" And SubInv = "" Then
            strSQL = "select OrgCode,MaterialNo,SubInventory,Locator,LotNo,MFG,MPN,Total from v_onhandmfgmpn where OrgCode='" + Org + "' and MaterialNo in (" + sMaterial + ")"
        ElseIf sMaterial <> "" And SubInv <> "" Then
            strSQL = "select OrgCode,MaterialNo,SubInventory,Locator,LotNo,MFG,MPN,Total from v_onhandmfgmpn where OrgCode='" + Org + "' and MaterialNo in (" + sMaterial + ") and SubInv='" + SubInv + "'"
        End If

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim DA As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(strSQL, myConn)
        Dim DS As DataSet = New DataSet("DS")
        Try
            DA.SelectCommand.CommandTimeout = TimeOut_M60
            DA.Fill(DS)
        Catch ex As Exception
            ErrorLogging("Report-OnHandMFGMPN", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
        Return DS

    End Function

    Public Function TIssueCompare(ByVal DJ As String) As DataSet
        Dim strSQL As String
        strSQL = "exec eTraceAddition.dbo.Trace_QueryIssueCompare '" + DJ + "'"
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim DA As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(strSQL, myConn)
        Dim DS As DataSet = New DataSet
        Try
            DA.SelectCommand.CommandTimeout = TimeOut_M60
            DA.Fill(DS, "tblIssueQuery")
        Catch ex As Exception
            ErrorLogging("Report-TIssueCompare", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
        Return DS
    End Function

    Public Function TMaterialTransfer(ByVal Org As String, ByVal MaterialNo As String, ByVal MPN As String) As DataSet
        Dim sSql As String = "exec SP_MaterialTransferRpt '" + Org + "','" + MaterialNo + "','" + MPN + "'"
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim DA As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(sSql, myConn)
        Dim DS As DataSet = New DataSet
        Try
            DA.SelectCommand.CommandTimeout = TimeOut_M60
            DA.Fill(DS, "tblMaterialTransfer")
        Catch ex As Exception
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
        Return DS
    End Function

    Public Function GeteTraceOH(ByVal Org As String, ByVal ItemNo As String) As DataSet

        GeteTraceOH = Nothing

        Using myconn As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            myconn.Open()
            Try
                Dim ds As New DataSet
                ds.Clear()

                Dim strsql As String = " select OrgCode, MaterialNo as ItemNo,Sloc as SubInv ,Manufacturer,Manufacturerpn,DateCode,Sum(QtyBaseUOM) as QTY "
                strsql = strsql & " from t_clmaster with(nolock) where Statuscode=1 and materialno='" & ItemNo & "' and OrgCode='" & Org & "' "
                strsql = strsql & "  group by OrgCode , materialno,Sloc ,Manufacturer,Manufacturerpn,DateCode"
                Dim ada As SqlDataAdapter = New SqlDataAdapter(strsql, myconn)
                ada.SelectCommand.CommandTimeout = TimeOut_M60
                ada.Fill(ds, "eTraceOH")

                GeteTraceOH = ds

                myconn.Close()

            Catch ex As Exception

                GeteTraceOH = Nothing
                myconn.Close()

            End Try

        End Using

    End Function

    Public Function StandardTime() As String
        Try
            StandardTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
        Catch ex As Exception
            StandardTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
            ErrorLogging("SFC-StadardTime", "", ex.Message & ex.Source, "E")
        Finally
        End Try
    End Function

    Public Function GetOHQTYWithMPNList(ByVal Org As String, ByVal MPNlist As String) As DataSet

        GetOHQTYWithMPNList = Nothing

        Dim strItem As String = ""
        Dim ItemArray() As String = MPNlist.Split(",")

        For i As Integer = 0 To ItemArray.Length - 1

            If String.IsNullOrEmpty(strItem) Then
                strItem = "  manufacturerpn ='" & ItemArray(i).ToString.Trim & "'  "
            Else
                strItem = strItem & " or manufacturerpn='" & ItemArray(i).ToString.Trim & "'  "
            End If

        Next

        strItem = " ( " & strItem & " )"

        Using myconn As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            myconn.Open()
            Try
                Dim ds As New DataSet
                ds.Clear()

                Dim strsql As String = " select OrgCode,MaterialNo as ItemNo, "

                'strsql = "SLOC as SubInv ," 
                strsql = strsql & " SLOC as SubInv ,"
                strsql = strsql & " storagebin as Locator ,"
                strsql = strsql & " RecDocNo as RT,"
                strsql = strsql & " Manufacturer,"
                strsql = strsql & " Manufacturerpn,"
                strsql = strsql & " sum(QtyBaseUom) as OH_QTY from t_clmaster with(nolock)"

                strsql = strsql & " where OrgCode='" & Org & "' and statuscode=1 and storagebin<>'' and " & strItem

                strsql = strsql & " group by OrgCode,MaterialNo, SLOC ,storagebin ,RecDocNo ,Manufacturer,manufacturerpn"
                strsql = strsql & " order by MaterialNo"

                Dim ada As SqlDataAdapter = New SqlDataAdapter(strsql, myconn)
                ada.SelectCommand.CommandTimeout = 1000 * 60 * 60
                ada.Fill(ds, "eTraceOHMPN")

                GetOHQTYWithMPNList = ds

                myconn.Close()

            Catch ex As Exception

                GetOHQTYWithMPNList = Nothing
                myconn.Close()

            End Try

        End Using

    End Function

    Public Function GeteTraceItemOHMPQ(ByVal Org As String, ByVal Itemlist As String) As DataSet

        GeteTraceItemOHMPQ = Nothing

        Dim strItem As String = ""
        Dim ItemArray() As String = Itemlist.Split(",")

        For i As Integer = 0 To ItemArray.Length - 1

            If String.IsNullOrEmpty(strItem) Then
                strItem = "  materialno ='" & ItemArray(i).ToString.Trim & "'  "
            Else
                strItem = strItem & " or materialno='" & ItemArray(i).ToString.Trim & "'  "
            End If

        Next

        strItem = " ( " & strItem & " )"

        Using myconn As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            myconn.Open()
            Try
                Dim ds As New DataSet
                ds.Clear()

                Dim strsql As String = ""

                strsql = " select "
                strsql = strsql & " orgcode,"
                strsql = strsql & " materialno ,  "
                strsql = strsql & " Count(CLID) as CLID_QTY ,"
                strsql = strsql & " Sum(QTYBaseUOM) as Onhand_QTY ,"
                strsql = strsql & " Max(QTYBaseUOM) as MPQ"
                strsql = strsql & " from t_clmaster with(nolock) "
                strsql = strsql & " where orgcode='" & Org & "' "
                strsql = strsql & " and right(sloc,3)  in ('MC1','SM1','RWK','CLI')"
                strsql = strsql & " and statuscode=1 and " & strItem
                strsql = strsql & " group by orgcode, materialno"

                Dim ada As SqlDataAdapter = New SqlDataAdapter(strsql, myconn)
                ada.SelectCommand.CommandTimeout = 1000 * 60 * 60
                ada.Fill(ds, "eTraceOHMPQ")

                GeteTraceItemOHMPQ = ds

                myconn.Close()

            Catch ex As Exception

                GeteTraceItemOHMPQ = Nothing
                myconn.Close()

            End Try

        End Using

    End Function

    Public Function GetConnectString(ByVal Type As String) As String
        If Type.ToUpper = "Onetoone".ToUpper Then
            Return System.Configuration.ConfigurationManager.AppSettings("eTraceOTOConnString")
        ElseIf Type.ToUpper = "eTraceVI".ToUpper Then
            Return System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString")
        ElseIf Type.ToUpper = "eTraceVII".ToUpper Then
            Return System.Configuration.ConfigurationManager.AppSettings("eTraceDBVIIConnString")
        End If
    End Function


#Region "IProAMLVSeTrace"

    Public Function IProAMLVSeTrace(ByVal strOrgCode As String, ByVal strSubINV As String) As DataSet

        IProAMLVSeTrace = New DataSet
        IProAMLVSeTrace.Clear()

        Dim strcondition As String = ""
        Dim SQLCMD As SqlCommand
        Dim myConn As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String = ""
        Dim ds As New DataSet
        Dim objReader As SqlDataReader

        strcondition = strcondition & "and " & "OrgCode ='" + strOrgCode.Trim() + "' "
        strcondition = strcondition & "and " & "SLoc ='" + strSubINV.Trim() + "' "

        strSQL = "Select CLID,RTLot,MaterialNo,OrgCode,SLoc,StorageBin,manufacturer as eTraceMFR, Manufacturerpn as eTraceMPN "
        strSQL = strSQL & " from T_CLMaster with (nolock) Where statuscode = 1 and qtybaseuom >0 and " & Mid(strcondition, 4)

        Try

            myConn.Open()
            SQLCMD = New SqlCommand(strSQL, myConn)
            SQLCMD.CommandTimeout = TimeOut_M60
            objReader = SQLCMD.ExecuteReader()

            Dim DTab As New DataTable
            DTab.Load(objReader)

            DTab.Columns.Add("iProMFR", Type.GetType("System.String"))
            DTab.Columns.Add("iProMPN", Type.GetType("System.String"))

            Dim tabResult As DataTable = DTab.Clone
            Dim iproData() As String

            If (DTab.Rows.Count > 0) Then

                Dim drow As DataRow
                Dim n As Integer = DTab.Rows.Count
                Dim k As Integer = 0
                ReDim iproData(n)
                For Each drow In DTab.Rows
                    iproData(k) = drow.Item("MaterialNo").ToString.Trim
                    k = k + 1
                Next

            Else
                Exit Function

            End If

            Dim ds1 As DataSet
            ds1 = GetAML(iproData)

            Dim ii As Integer = ds1.Tables(0).Rows.Count

            Dim eTraceRow As DataRow
            Dim eTraceItemNo As String
            Dim i As Integer = 0
            Dim j As Integer = 0

            For Each eTraceRow In DTab.Rows

                eTraceItemNo = eTraceRow.Item("MaterialNo").ToString.Trim
                Dim findrows() As DataRow = ds1.Tables(0).Select("MaterialNo = '" & eTraceItemNo & "' and AMLStatus='Active' ")
                Dim striproItem As String = ""
                Dim striproMFR As String = ""
                Dim BLCheck As Boolean = False

                If findrows.Length > 0 Then

                    For j = 0 To findrows.Length - 1
                        If DTab.Rows(i)("eTraceMFR").ToString.Trim = findrows(j).Item("MFR").ToString.Trim And DTab.Rows(i)("eTraceMPN").ToString.Trim = findrows(j).Item("MPN").ToString.Trim Then
                            BLCheck = True
                            Exit For
                        End If
                    Next

                End If

                If findrows.Length > 0 And BLCheck = False Then '一个eTrace MFR对应多个iPro MFR

                    For j = 0 To findrows.Length - 1

                        Dim Resultrow As DataRow
                        Resultrow = tabResult.NewRow

                        Resultrow.Item("OrgCode") = DTab.Rows(i)("OrgCode")
                        Resultrow.Item("CLID") = DTab.Rows(i)("CLID")
                        Resultrow.Item("MaterialNo") = DTab.Rows(i)("MaterialNo")
                        Resultrow.Item("SLoc") = DTab.Rows(i)("SLoc")
                        Resultrow.Item("StorageBin") = DTab.Rows(i)("StorageBin")
                        Resultrow.Item("RTLot") = DTab.Rows(i)("RTLot")
                        Resultrow.Item("eTraceMFR") = DTab.Rows(i)("eTraceMFR")
                        Resultrow.Item("eTraceMPN") = DTab.Rows(i)("eTraceMPN")
                        Resultrow.Item("iProMFR") = findrows(j).Item("MFR")
                        Resultrow.Item("iProMPN") = findrows(j).Item("MPN")

                        tabResult.Rows.Add(Resultrow)

                    Next

                End If

                i += 1

            Next

            IProAMLVSeTrace.Tables.Add(tabResult)

        Catch ex As Exception
            ErrorLogging("Report-IProAMLVSeTrace", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function IProAMLVSeTrace2(ByVal strOrgCode As String, ByVal strSubINV As String, ByVal rtDateFrom As String, ByVal rtDateTo As String, ByVal amlStatus As String, ByVal clidStatus As String) As DataSet

        IProAMLVSeTrace2 = New DataSet
        IProAMLVSeTrace2.Clear()

        Dim strcondition As String = ""
        Dim SQLCMD As SqlCommand
        Dim myConn As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim strSQL As String = ""
        Dim ds As New DataSet
        Dim objReader As SqlDataReader

        strcondition = strcondition & "and " & "OrgCode ='" + strOrgCode.Trim() + "' "

        If strSubINV.Trim() <> "" Then
            strcondition = strcondition & "and " & "SLoc ='" + strSubINV.Trim() + "' "
        End If

        If rtDateTo.Trim() <> "" And rtDateFrom.Trim() <> "" Then
            strcondition = strcondition & "and " & "RecDate between '" + rtDateFrom.Trim() + " 00:00:00 AM' and '" & rtDateTo.Trim() & " 00:00:00 AM' "
        End If

        If amlStatus <> "" Then
            strcondition = strcondition & "and " & "QMLStatus ='" + amlStatus.Trim() + "' "
        End If

        If clidStatus <> "0&1" Then
            strcondition = strcondition & "and " & "StatusCode ='" + clidStatus.Trim() + "' "
        End If

        strSQL = "Select OrgCode,CLID,MaterialNo,MaterialDesc,QtyBaseUOM,BaseUOM,CreatedBy,ExpDate,RecDocNo,PurOrdNo,DeliveryType,VendorID,VendorName,VendorPN,InvoiceNo,HeaderText,Operator,StockType,manufacturer as eTraceMFR,Manufacturerpn as eTraceMPN,QMLStatus,StatusCode "
        strSQL = strSQL & " from T_CLMaster with (nolock) Where 1 = 1 and qtybaseuom > 0 and " & Mid(strcondition, 4)

        Try

            myConn.Open()
            SQLCMD = New SqlCommand(strSQL, myConn)
            SQLCMD.CommandTimeout = TimeOut_M60
            objReader = SQLCMD.ExecuteReader()

            Dim DTab As New DataTable
            DTab.Load(objReader)

            DTab.Columns.Add("iProMFR", Type.GetType("System.String"))
            DTab.Columns.Add("iProMPN", Type.GetType("System.String"))
            DTab.Columns.Add("iProStatus", Type.GetType("System.String"))

            Dim tabResult As DataTable = DTab.Clone
            Dim iproData() As String

            If (DTab.Rows.Count > 0) Then

                Dim drow As DataRow
                Dim n As Integer = DTab.Rows.Count
                Dim k As Integer = 0
                ReDim iproData(n)
                For Each drow In DTab.Rows
                    iproData(k) = drow.Item("MaterialNo").ToString.Trim
                    k = k + 1
                Next

            Else
                Exit Function

            End If

            Dim t As New ArrayList()

            Dim num As Integer
            For num = 0 To iproData.Length - 1
                If num = 0 And Not iproData Is Nothing Then
                    t.Add(iproData(num))
                Else
                    If Not t.Contains(iproData(num)) And Not iproData(num) Is Nothing Then
                        t.Add(iproData(num))
                    End If
                End If
            Next

            If t.Count = 0 Then
                Exit Function
            End If

            Dim statusCode As String
            statusCode = "ACTIVE,EMBEDDED POWER DISCONTINUED,TEMP RESTRICTED"

            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Sqlstr = String.Format("Select Value from T_Config where ConfigID = 'REC003'")
                statusCode = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
            End Using

            statusCode = " '" & statusCode.Replace(",", "','") & "' "

            Dim ds1 As DataSet
            ds1 = GetAML(t.ToArray(GetType(String)))

            Dim ii As Integer = ds1.Tables(0).Rows.Count

            Dim eTraceRow As DataRow
            Dim eTraceItemNo As String
            Dim i As Integer = 0
            Dim j As Integer = 0

            For Each eTraceRow In DTab.Rows

                eTraceItemNo = eTraceRow.Item("MaterialNo").ToString.Trim
                Dim findrows() As DataRow = ds1.Tables(0).Select("MaterialNo = '" & eTraceItemNo & "' and AMLStatus in (" & statusCode & ") ")
                Dim striproItem As String = ""
                Dim striproMFR As String = ""
                Dim BLCheck As Boolean = False

                If findrows.Length > 0 Then

                    For j = 0 To findrows.Length - 1
                        If DTab.Rows(i)("eTraceMFR").ToString.Trim = findrows(j).Item("MFR").ToString.Trim And DTab.Rows(i)("eTraceMPN").ToString.Trim = findrows(j).Item("MPN").ToString.Trim Then
                            BLCheck = True
                            Exit For
                        End If
                    Next

                End If

                If findrows.Length > 0 And BLCheck = False Then '??eTrace MFR????iPro MFR

                    For j = 0 To findrows.Length - 1

                        Dim Resultrow As DataRow
                        Resultrow = tabResult.NewRow

                        Resultrow.Item("OrgCode") = DTab.Rows(i)("OrgCode")
                        Resultrow.Item("CLID") = DTab.Rows(i)("CLID")
                        Resultrow.Item("MaterialNo") = DTab.Rows(i)("MaterialNo")
                        Resultrow.Item("MaterialDesc") = DTab.Rows(i)("MaterialDesc")
                        Resultrow.Item("QtyBaseUOM") = DTab.Rows(i)("QtyBaseUOM")
                        Resultrow.Item("BaseUOM") = DTab.Rows(i)("BaseUOM")
                        Resultrow.Item("CreatedBy") = DTab.Rows(i)("CreatedBy")
                        Resultrow.Item("ExpDate") = DTab.Rows(i)("ExpDate")
                        Resultrow.Item("RecDocNo") = DTab.Rows(i)("RecDocNo")
                        Resultrow.Item("PurOrdNo") = DTab.Rows(i)("PurOrdNo")
                        Resultrow.Item("DeliveryType") = DTab.Rows(i)("DeliveryType")
                        Resultrow.Item("VendorID") = DTab.Rows(i)("VendorID")
                        Resultrow.Item("VendorName") = DTab.Rows(i)("VendorName")
                        Resultrow.Item("VendorPN") = DTab.Rows(i)("VendorPN")
                        Resultrow.Item("InvoiceNo") = DTab.Rows(i)("InvoiceNo")
                        Resultrow.Item("HeaderText") = DTab.Rows(i)("HeaderText")
                        Resultrow.Item("Operator") = DTab.Rows(i)("Operator")
                        Resultrow.Item("StockType") = DTab.Rows(i)("StockType")
                        Resultrow.Item("QMLStatus") = DTab.Rows(i)("QMLStatus")
                        Resultrow.Item("StatusCode") = DTab.Rows(i)("StatusCode")
                        Resultrow.Item("eTraceMFR") = DTab.Rows(i)("eTraceMFR")
                        Resultrow.Item("eTraceMPN") = DTab.Rows(i)("eTraceMPN")
                        Resultrow.Item("iProMFR") = findrows(j).Item("MFR")
                        Resultrow.Item("iProMPN") = findrows(j).Item("MPN")
                        Resultrow.Item("iProStatus") = findrows(j).Item("AMLStatus")

                        tabResult.Rows.Add(Resultrow)

                    Next

                End If

                i += 1

            Next

            IProAMLVSeTrace2.Tables.Add(tabResult)

        Catch ex As Exception
            ErrorLogging("Report-IProAMLVSeTrace2", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

#End Region


    'External Function Called by SMT and other detartments
#Region "External Function"

    Public Function GetMaterialInfoByCLID(ByVal CLID As String) As DataSet
        Dim myConn As SqlClient.SqlConnection
        Dim Sqlstr1 As String
        Dim ds As New DataSet
        Try
            myConn = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceConnectionStringRpt"))
            Sqlstr1 = String.Format("sp_getMaterialInfoByCLID")
            Dim _SCmd As New SqlClient.SqlCommand(Sqlstr1, myConn)
            Dim _SApt As New SqlClient.SqlDataAdapter(_SCmd)
            Dim dt As New DataTable

            _SCmd.CommandType = CommandType.StoredProcedure
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Parameters.Add("@CLID", SqlDbType.VarChar, 200, "CLID")
            _SCmd.Parameters("@CLID").Value = CLID

            _SApt.Fill(dt)
            myConn.Close()
            ds.Tables.Add(dt)

            Return ds
        Catch ex As Exception
            'ErrorLogging("ExecuteDataTable", "", Sqlstr1 & ", with error: " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try


        'Dim ds As New DataSet
        'Dim Sqlstr1 As String
        'Using da As DataAccess = GetDataAccess()
        '    Try
        '        Sqlstr1 = String.Format("exec sp_getMaterialInfoByCLID '{0}'", CLID)
        '        ds = da.ExecuteDataSet(Sqlstr1)
        '    Catch ex As Exception
        '        ErrorLogging("GetMaterialInfoByCLID", "", ex.Message & ex.Source, "E")
        '    End Try
        '    Return ds
        'End Using
    End Function

    Public Function GetMaterialToXMLByCLID(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As String

        GetMaterialToXMLByCLID = "Fail to create xml file"
        Dim xmlfilename As String
        Dim ds As New DataSet()
        Dim Sqlstr1 As String
        Dim i As Integer
        Dim dr As DataRow
        xmlfilename = Guid.NewGuid.ToString + ".xml"


        Dim myConn As SqlClient.SqlConnection
        'Try




        '  Using da As DataAccess = GetDataAccess()
        Try
            'Sqlstr1 = String.Format("select MaterialNo AS PartNumber, CLID as AGI, LotNo AS LotNumber, QtyBaseUOM AS Quantity, (StatusCode + 1) % 2 AS StatusCode from T_CLmaster where createdon between '{0}' and '{1}' ", From_MR, To_MR)
            'Sqlstr1 = String.Format("exec sp_getMaterialInfosBydate '{0}' ,'{1}','{2}'", From_MR, To_MR, pn)
            'ds = da.ExecuteDataSet(Sqlstr1, "InventoryItemData")
            'ds.DataSetName = "InventoryData"
            myConn = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceConnectionStringRpt"))
            Sqlstr1 = String.Format("sp_getMaterialInfosBydate")
            Dim _SCmd As New SqlClient.SqlCommand(Sqlstr1, myConn)
            Dim _SApt As New SqlClient.SqlDataAdapter(_SCmd)
            Dim dt As New DataTable

            _SCmd.CommandType = CommandType.StoredProcedure
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Parameters.Add("@From_MR", SqlDbType.SmallDateTime, 4, "From_MR")
            _SCmd.Parameters("@From_MR").Value = From_MR

            _SCmd.Parameters.Add("@To_MR", SqlDbType.SmallDateTime, 4, "To_MR")
            _SCmd.Parameters("@To_MR").Value = To_MR

            _SCmd.Parameters.Add("@pn", SqlDbType.VarChar, 50, "pn")
            _SCmd.Parameters("@pn").Value = pn

            _SApt.Fill(dt)
            myConn.Close()
            ds.Tables.Add(dt)

            If (ds IsNot Nothing And ds.Tables(0).Rows.Count > 0) Then



                Dim myTW As New XmlTextWriter(ConfigurationSettings.AppSettings.Item("XMLFolder") + xmlfilename, Nothing)

                myTW.WriteStartDocument()
                'myTW.Formatting = Formatting.Indented


                myTW.WriteStartElement("InventoryData")
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dr = ds.Tables(0).Rows(i)
                    myTW.WriteStartElement("InventoryItemData")
                    myTW.WriteAttributeString("PartNumber", dr(0).ToString)
                    myTW.WriteAttributeString("AGI", dr(1).ToString)
                    myTW.WriteAttributeString("LotNumber", dr(2).ToString)
                    myTW.WriteAttributeString("Quantity", dr(3).ToString)
                    myTW.WriteAttributeString("StatusCode", dr(4).ToString)
                    myTW.WriteFullEndElement()
                Next
                myTW.WriteFullEndElement()
                myTW.WriteEndDocument()
                myTW.Flush()
                myTW.Close()

                'ds.WriteXml(ConfigurationSettings.AppSettings.Item("XMLFolder") + xmlfilename)
                GetMaterialToXMLByCLID = xmlfilename
            End If
        Catch ex As Exception
            GetMaterialToXMLByCLID = GetMaterialToXMLByCLID + " " + ex.Message
            ErrorLogging("GetMaterialToXMLByCLID", "", ex.Message & ex.Source, "E")

        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try



        '  End Using
    End Function

    Public Function GetMaterialToXMLByCLID2(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As String

        GetMaterialToXMLByCLID2 = "Fail to create xml file"
        Dim xmlfilename As String
        Dim ds As New DataSet()
        Dim Sqlstr1 As String
        Dim i As Integer
        Dim dr As DataRow
        Dim myConn As SqlClient.SqlConnection

        Try
            pn = FixNull(pn)
            xmlfilename = Guid.NewGuid.ToString + ".xml"

            myConn = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceConnectionStringRpt"))
            Sqlstr1 = String.Format("sp_getMaterialInfosBydate2")
            Dim _SCmd As New SqlClient.SqlCommand(Sqlstr1, myConn)
            Dim _SApt As New SqlClient.SqlDataAdapter(_SCmd)
            Dim dt As New DataTable
            dt.TableName = "InventoryItemData"

            _SCmd.CommandType = CommandType.StoredProcedure
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Parameters.Add("@From_MR", SqlDbType.SmallDateTime, 4, "From_MR")
            _SCmd.Parameters("@From_MR").Value = From_MR

            _SCmd.Parameters.Add("@To_MR", SqlDbType.SmallDateTime, 4, "To_MR")
            _SCmd.Parameters("@To_MR").Value = To_MR

            _SCmd.Parameters.Add("@pn", SqlDbType.VarChar, 50, "pn")
            _SCmd.Parameters("@pn").Value = pn

            _SApt.Fill(dt)
            myConn.Close()
            ds.Tables.Add(dt)

            Dim Sqlstr As String
            Dim da As DataAccess = GetDataAccess()

            Dim RTNCLIDs As New DataSet
            Sqlstr = String.Format("exec sp_getMaterialInfosBydate2  '{0}','{1}','{2}' ", From_MR, To_MR, pn)  'the SP in eTraccedbzs is different from rpt01
            RTNCLIDs = da.ExecuteDataSet(Sqlstr, "InventoryItemData")

            If RTNCLIDs Is Nothing OrElse RTNCLIDs.Tables.Count = 0 Then
            ElseIf RTNCLIDs.Tables(0).Rows.Count > 0 Then
                ds.Merge(RTNCLIDs)
            End If

            If (ds IsNot Nothing And ds.Tables(0).Rows.Count > 0) Then
                ds.DataSetName = "InventoryData"
                ds.WriteXml(ConfigurationSettings.AppSettings.Item("XMLFolder2") + xmlfilename)
                GetMaterialToXMLByCLID2 = xmlfilename
            End If

            'Record the FromDate and ToDate as part of File name  -- 01/24/2017
            Dim FileDateStr As String = "_" + Format(From_MR, "yyyy-MM-dd H.mm.ss").Replace(" ", "t") + "_" + Format(To_MR, "yyyy-MM-dd H.mm.ss").Replace(" ", "t")

            'Write xml file to another folder for Siplace use -- 9/18/2015
            'Dim Siplacexml As String = Format(Now(), "yyyy-MM-dd H.mm.ss").Replace(" ", "T") + ".xml"
            Dim Siplacexml As String = Format(Now(), "yyyy-MM-dd H.mm.ss").Replace(" ", "T") + FileDateStr + ".xml"

            '<Packaging IdPackaging="ZS000000000024690304" ComponentName="131-000334-0000" ComponentBarcode="" ComponentBarcodeFragment="131-000334-0000" Quantity="500" Manufacturer="INFINEON" 
            'ManufactureDate="2015-07-06T04:00:00+08:00" ExpiryDate="2017-07-05T00:00:00+08:00" BatchId="ZS000000000024690304" MsdLevel="1" ManufacturerPartNumber="IPP60R199CP" RoHS="Y" Batch2="PF524273W06"/>
            If (ds IsNot Nothing And ds.Tables(0).Rows.Count > 0) Then
                Dim siplaceTW As New XmlTextWriter(ConfigurationSettings.AppSettings.Item("SiplaceXMLFolder") + Siplacexml, Nothing)
                siplaceTW.WriteStartDocument()
                siplaceTW.WriteStartElement("PackagingList")
                siplaceTW.WriteAttributeString("ImportVersion", "3.0.1.1")
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    dr = ds.Tables(0).Rows(i)
                    siplaceTW.WriteStartElement("Packaging")
                    siplaceTW.WriteAttributeString("IdPackaging", dr("UID").ToString)
                    siplaceTW.WriteAttributeString("ComponentName", dr("PartNumber").ToString)

                    'siplaceTW.WriteAttributeString("ComponentBarcode", "")
                    'siplaceTW.WriteAttributeString("ComponentBarcodeFragment", dr("PartNumber").ToString)
                    siplaceTW.WriteAttributeString("ComponentBarcodeFragment", "")
                    siplaceTW.WriteAttributeString("ComponentBarcode", dr("PartNumber").ToString)

                    siplaceTW.WriteAttributeString("Quantity", dr("Quantity").ToString)
                    siplaceTW.WriteAttributeString("Manufacturer", dr("Supplier").ToString)

                    'siplaceTW.WriteAttributeString("ManufactureDate", dr("ManufacturedDate").ToString.Replace("/", "-").Replace(" ", "T") + "+08:00")
                    'siplaceTW.WriteAttributeString("ExpiryDate", dr("ExpirationDate").ToString.Replace("/", "-").Replace(" ", "T") + "+08:00")

                    If ChangeDateFormat(dr("ManufacturedDate").ToString) <> "" Then
                        siplaceTW.WriteAttributeString("ManufactureDate", ChangeDateFormat(dr("ManufacturedDate").ToString))
                    End If

                    If ChangeDateFormat(dr("ExpirationDate").ToString) <> "" Then
                        siplaceTW.WriteAttributeString("ExpiryDate", ChangeDateFormat(dr("ExpirationDate").ToString))
                    End If

                    siplaceTW.WriteAttributeString("BatchId", dr("UID").ToString)

                    If dr("JedecLevel").ToString <> "" Then
                        siplaceTW.WriteAttributeString("MsdLevel", dr("JedecLevel").ToString)
                    End If

                    'siplaceTW.WriteAttributeString("ManufacturerPartNumber", dr("MfgPartNumber").ToString)
                    siplaceTW.WriteAttributeString("ManufacturerPartNumber", Left(dr("MfgPartNumber").ToString, 40))

                    If dr("RoHS").ToString.ToUpper = "TRUE" Then
                        siplaceTW.WriteAttributeString("RoHS", "Y")
                    Else
                        siplaceTW.WriteAttributeString("RoHS", "N")
                    End If
                    siplaceTW.WriteAttributeString("Batch2", Left(dr("LotNumber").ToString, 40))

                    siplaceTW.WriteFullEndElement()
                Next
                siplaceTW.WriteFullEndElement()
                siplaceTW.WriteEndDocument()
                siplaceTW.Flush()
                siplaceTW.Close()
            End If


        Catch ex As Exception
            GetMaterialToXMLByCLID2 = GetMaterialToXMLByCLID2 + " " + ex.Message
            ErrorLogging("GetMaterialToXMLByCLID2", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Private Function ChangeDateFormat(ByVal oriDate As String) As String
        ChangeDateFormat = ""

        Try
            If oriDate <> "" Then
                ChangeDateFormat = oriDate.ToString.Replace("/", "-").Replace(" ", "T") + "+08:00"
            End If
        Catch ex As Exception
        End Try

    End Function

    Public Function GetDJToXML() As String

        GetDJToXML = "Fail to create xml file"

        Dim xmlfilename As String
        Dim ds As New DataSet()

        Try

            Dim Sqlstr As String
            Dim da As DataAccess = GetDataAccess()

            Sqlstr = String.Format("exec sp_GetDJInfoforSMT ")
            ds = da.ExecuteDataSet(Sqlstr, "BatchData")

            If (ds IsNot Nothing And ds.Tables(0).Rows.Count > 0) Then
                xmlfilename = Guid.NewGuid.ToString + "_Batches.xml"
                ds.DataSetName = "Batches"
                ds.WriteXml(ConfigurationSettings.AppSettings.Item("XMLFolder3") + xmlfilename)
                GetDJToXML = xmlfilename
            End If

        Catch ex As Exception
            GetDJToXML = GetDJToXML + " " + ex.Message
            ErrorLogging("GetDJToXML", "", ex.Message & ex.Source, "E")
        End Try

    End Function

    Public Function GetMaterialByCLIDToDataSet(ByVal From_MR As DateTime, ByVal To_MR As DateTime, ByVal pn As String) As DataSet

        'GetMaterialToXMLByCLID = "Fail to create xml file"
        'Dim xmlfilename As String
        Dim ds As New DataSet()
        Dim Sqlstr1 As String
        Dim i As Integer
        Dim dr As DataRow
        'xmlfilename = Guid.NewGuid.ToString + ".xml"

        Dim myConn As SqlClient.SqlConnection
        'Try

        '  Using da As DataAccess = GetDataAccess()
        Try
            'Sqlstr1 = String.Format("select MaterialNo AS PartNumber, CLID as AGI, LotNo AS LotNumber, QtyBaseUOM AS Quantity, (StatusCode + 1) % 2 AS StatusCode from T_CLmaster where createdon between '{0}' and '{1}' ", From_MR, To_MR)
            'Sqlstr1 = String.Format("exec sp_getMaterialInfosBydate '{0}' ,'{1}','{2}'", From_MR, To_MR, pn)
            'ds = da.ExecuteDataSet(Sqlstr1, "InventoryItemData")
            'ds.DataSetName = "InventoryData"
            myConn = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceConnectionStringRpt"))
            Sqlstr1 = String.Format("sp_getMaterialInfosBydate")
            Dim _SCmd As New SqlClient.SqlCommand(Sqlstr1, myConn)
            Dim _SApt As New SqlClient.SqlDataAdapter(_SCmd)
            Dim dt As New DataTable

            _SCmd.CommandType = CommandType.StoredProcedure
            _SCmd.CommandTimeout = TimeOut_M30
            _SCmd.Parameters.Add("@From_MR", SqlDbType.SmallDateTime, 4, "From_MR")
            _SCmd.Parameters("@From_MR").Value = From_MR

            _SCmd.Parameters.Add("@To_MR", SqlDbType.SmallDateTime, 4, "To_MR")
            _SCmd.Parameters("@To_MR").Value = To_MR

            _SCmd.Parameters.Add("@pn", SqlDbType.VarChar, 50, "pn")
            _SCmd.Parameters("@pn").Value = pn

            _SApt.Fill(dt)
            myConn.Close()
            ds.Tables.Add(dt)

            Return ds
        Catch ex As Exception
            'GetMaterialToXMLByCLID = GetMaterialToXMLByCLID + " " + ex.Message
            ErrorLogging("GetMaterialByCLIDToDataSet", "", ex.Message & ex.Source, "E")
            Return Nothing
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        Return ds

        '  End Using
    End Function

    Public Function GetReturnCLIDByDate(ByVal From_RTN As DateTime, ByVal To_RTN As DateTime) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs As DataSet = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetReturnCLIDByDate  '{0}','{1}' ", From_RTN, To_RTN)
                myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDs")
                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("GetReturnCLIDByDate", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

#End Region

End Class
