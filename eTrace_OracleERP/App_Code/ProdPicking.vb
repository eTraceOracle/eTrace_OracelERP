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
Imports System.Math

Public Structure MOData
    Public MO As String
    Public MOType As String
    Public OrgID As String
    Public ItemID As String
    Public MOHeaderID As String
    Public SpecialItem As String
    Public SpecialSubInv As String
    Public PickFlag As String           '0 = "Normal DJ/MO Confirm"; 1 = "SMT Special MO Confirm"
    Public Items As DataSet
End Structure

Public Structure SMTData
    Public UserProdLine As String
    Public OrgCode As String
    Public EventID As String
    Public EventType As String
    Public MONo As String
    Public User As String
    Public CLID As String
    Public ActionType As String
    Public CheckDock As String
    Public RtnMsg As String
    Public dsItem As DataSet
    Public BoxID As String
End Structure

Public Structure ProcessLabel
    Public ProcessLabelID As String
    Public DJSet As String
    Public Model As String
    Public MaterialNo As String
    Public ReqQty As String
    Public UsageQty As String
    Public FormingCode As String
    Public RefDesignator As String
    Public WorkStation As String
    Public AutoManual As String
End Structure

Public Structure PickingLabel
    Public OrgCode As String
    Public EventID As String
    Public EventDate As String                        'EventID Closed Date
    Public MONo As String
    Public DJ As String
    Public Job As String
    Public PCBA As String
    Public DJRev As String
    Public DJQty As String
    Public ProdLine As String
    Public Floor As String
    Public CLIDCount As String
    Public PanelSide As String
    Public GroupFlag As String
End Structure

Public Structure LDPickingLabel
    Public OrgCode As String
    Public EventID As String
    Public EventDate As String                        'EventID Closed Date
    Public MONo As String
    Public DJ As String
    Public Job As String
    Public PCBA As String
    Public DJRev As String
    Public DJQty As String
    Public ProdLine1 As String
    Public ProdLine2 As String
    Public ProdLine3 As String
    Public Floor As String
    Public CLIDCount As String
End Structure

Public Structure LDSpareLabel
    Public OrgCode As String
    Public EventID As String
    Public MONo As String
    Public DJ As String
    Public Job As String
    Public PCBA As String
    Public DJQty As String
    Public GroupBagID As String                          'GroupBag
    Public ProdLine As String
    Public Panel As String
End Structure

Public Class ProdPicking
    Inherits PublicFunction
    Private ProcessLabelFile As String = "D:\eTrace\ProcessLabel.lab"
    Private PickingLabelFile As String = "D:\eTrace\PickingLabel.lab"
    Private LDPickingLabelFile As String = "D:\eTrace\LDPickingLabel.lab"
    Private LDSpareLabelFile As String = "D:\eTrace\LDSpareLabel.lab"

#Region "ProdPicking"

    Public Function CheckCHDJ(ByVal DJ As String, ByVal CLID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim myStatus As String = ""

            Try
                'Read DJ / Item Status from HuaWei Web Service 
                Dim myCHWS As CHWS.CollectDataToQDB = New CHWS.CollectDataToQDB

                If CLID = "" Then
                    myStatus = myCHWS.DJStatus(DJ)                                                          ' DJStatus = "0"  means OK
                Else
                    myStatus = myCHWS.ItemRelationStatus(DJ, CLID)                              ' CHItem = "0"  means OK
                End If

                If myStatus = "0" Then
                    CheckCHDJ = myStatus
                Else
                    CheckCHDJ = "Error"
                    ErrorLogging("ProdPicking-CheckCHDJ1", "", "DJ: " & DJ & "; " & myStatus, "I")
                End If

            Catch ex As Exception
                ErrorLogging("ProdPicking-CheckCHDJ", "", "DJ: " & DJ & "; " & ex.Message & ex.Source, "E")
                Return "Error"
            End Try
        End Using

    End Function

    Public Function ReadCLIDs(ByVal CLID As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_ReadLooseCLIDs '{0}', '{1}' ", LoginData.OrgCode, CLID)
                Return da.ExecuteDataSet(Sqlstr, "CLIDs")

            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadCLIDs", LoginData.User.ToUpper, "CLID: " & CLID & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function ReadCLIDData(ByVal CLID As String, ByVal TransactionType As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                Dim ds As DataSet = New DataSet()
                Dim CHID As String = CLID
                Dim CHItem As String = ""

                If (Mid(CLID, 3, 1) = "B" AndAlso Len(CLID) = 20) OrElse Mid(CLID, 1, 1) = "P" OrElse Mid(CLID, 1, 1) = "1" Then 'Box ID
                    'Read One CLID from table T_CLMaster first
                    Sqlstr = String.Format("Select TOP (1) CLID from T_CLMaster with (nolock) where BoxID = '{0}' and OrgCode='{1}' and (SLOC<>'' or not (SLOC IS NULL))", CLID, LoginData.OrgCode)
                    CHID = Convert.ToString(da.ExecuteScalar(Sqlstr))

                    Sqlstr = String.Format("Select OrgCode, '' as CLID, MaterialNo,RTLot,Sum(QtyBaseUOM) As QtyBaseUOM,BaseUOM,StatusCode,SLOC,StorageBin,Manufacturer,ManufacturerPN,MaterialRevision,ExpDate,BoxID,SONo,SOLine,CountryOfOrigin, CHItem ='' from T_CLMaster with (nolock) " _
                             & "group by OrgCode, MaterialNo,RTLot,BaseUOM,StatusCode,SLOC,StorageBin,Manufacturer,ManufacturerPN,MaterialRevision,ExpDate,BoxID,SONo,SOLine,CountryOfOrigin having BoxID = '{0}' and OrgCode='{1}' and (SLOC<>'' or not (SLOC IS NULL))", CLID, LoginData.OrgCode)
                Else
                    Sqlstr = String.Format("Select OrgCode,CLID,MaterialNo,RTLot,QtyBaseUOM,BaseUOM,StatusCode,SLOC,StorageBin,Manufacturer,ManufacturerPN,MaterialRevision,ExpDate,BoxID,SONo,SOLine,CountryOfOrigin, CHItem ='' from T_CLMaster with (nolock) where CLID = '{0}' and StatusCode = '1' and OrgCode='{1}' and (SLOC<>'' or not (SLOC IS NULL))", CLID, LoginData.OrgCode)
                End If
                ds = da.ExecuteDataSet(Sqlstr, "CLIDs")

                If TransactionType.Contains("/") Then                  'Record CH1 DJ here for CH1 Item Status checking
                    Dim CHDJ As String = Mid(TransactionType, 4)
                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    ElseIf ds.Tables(0).Rows.Count > 0 Then
                        'Read Item Status from HuaWei Web Service 
                        CHItem = CheckCHDJ(CHDJ, CHID)                         ' CHItem = "0"  means OK

                        ds.Tables(0).Rows(0)("CHItem") = CHItem
                        ds.Tables(0).AcceptChanges()
                    End If
                End If

                Return ds

            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadCLIDData", LoginData.User.ToUpper, "CLID: " & CLID & " with TransactionType: " & TransactionType & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function GetMOFromDJ(ByVal DJ As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda_s As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("MOrderList")
                oda_s.SelectCommand.CommandType = CommandType.StoredProcedure
                oda_s.SelectCommand.CommandText = "apps.XXETR_wip_pkg.get_morders"
                oda_s.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
                oda_s.SelectCommand.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 50)).Value = DJ
                oda_s.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda_s.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oda_s.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output
                oda_s.SelectCommand.Connection.Open()
                oda_s.Fill(ds, "MOrderList")
                oda_s.SelectCommand.Connection.Close()

                Return ds
            Catch ex As Exception
                ErrorLogging("ProdPicking-GetMOFromDJ", LoginData.User.ToUpper, "DJ: " & DJ & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            Finally
                If oda_s.SelectCommand.Connection.State <> ConnectionState.Closed Then oda_s.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function

    Public Function GetDJMOLines(ByVal DJ As String, ByVal MO As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            GetDJMOLines = New DataSet

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Dim Msg As String = ""
            Dim MaterialNo As String = ""
            If MO <> "" AndAlso MO.Contains("/") Then
                Dim ArryMO() As String = Split(MO, "/")
                MO = ArryMO(0)
                MaterialNo = ArryMO(1)
            End If

            If DJ <> "" Then Msg = "DJ: " & DJ & "; "
            If MO <> "" Then Msg = "MO: " & MO & "; "

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("MatlList")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.XXETR_wip_pkg.get_djmo_lines"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 50)).Value = DJ
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_move_order", OracleType.VarChar, 50)).Value = MO

                If MaterialNo <> "" Then
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_num", OracleType.VarChar, 50)).Value = MaterialNo
                End If

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "MatlList")
                oda.SelectCommand.Connection.Close()

                Dim CHDJ As String = DJ
                If ds.Tables("MatlList").Columns.Count > 1 Then
                    'If users already entered DJ Number, then need to remove other DJs if there exists 
                    If DJ <> "" AndAlso MO <> "" Then
                        Dim DJdr() As DataRow
                        DJdr = ds.Tables(0).Select("DJ_NAME <> '" & DJ & "'")
                        If DJdr.Length > 0 Then
                            Dim i As Integer
                            For i = 0 To DJdr.Length - 1
                                DJdr(i).Delete()
                            Next
                            ds.Tables(0).AcceptChanges()
                        End If
                    End If

                    'If DJ is blank, then read DJ Number from the first now of Oracle data table here
                    If DJ = "" AndAlso ds.Tables(0).Rows.Count > 0 Then
                        CHDJ = ds.Tables(0).Rows(0)("DJ_NAME").ToString
                    End If
                End If
                GetDJMOLines = ds

                ds = New DataSet
                Dim Sqlstr As String

                Sqlstr = String.Format("Select * from T_Config with (nolock) where ConfigID = 'PIK009' or ConfigID = 'MOC001' or ConfigID = 'WMS001' or ConfigID = 'CLID014' ")
                ds = da.ExecuteDataSet(Sqlstr, "Config")

                Dim DR() As DataRow
                Dim CHFlag As String = ""
                Dim DJStatus As String = ""
                DR = ds.Tables(0).Select(" ConfigID = 'CLID014' ")
                If DR.Length > 0 Then
                    CHFlag = DR(0)("Value").ToString

                    If CHFlag = "YES" AndAlso CHDJ <> "" Then
                        'Read DJ Status from HuaWei Web Service 
                        DJStatus = CheckCHDJ(CHDJ, "")                         ' DJStatus = "0"  means OK

                        DR(0)("Exempt") = DJStatus
                        DR(0).AcceptChanges()
                    End If
                End If

                GetDJMOLines.Merge(ds.Tables("Config"))

                Return GetDJMOLines

            Catch ex As Exception
                ErrorLogging("ProdPicking-GetDJMOLines", LoginData.User.ToUpper, Msg & ex.Message & ex.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using

    End Function

    'Public Function MOrderPost(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
    '    Using da As DataAccess = GetDataAccess()

    '        MOrderPost = ""

    '        Dim MO As String = myMOData.MO
    '        Dim dtPick As New DataTable
    '        Dim dsItem As DataSet = New DataSet

    '        Dim myDR As DataRow
    '        Dim dtModel As New DataTable
    '        dtModel.Columns.Add(New Data.DataColumn("DJ", System.Type.GetType("System.String")))
    '        dtModel.Columns.Add(New Data.DataColumn("Model", System.Type.GetType("System.String")))

    '        Dim Msg As String = "MO: " & myMOData.MO & "; "

    '        Try

    '            dsItem.DataSetName = "Items"
    '            dsItem.Merge(myMOData.Items.Tables(1))

    '            dtPick = myMOData.Items.Tables(0).Copy

    '            'Check if exists CLID's Status is 0, prevent user from slitting CLID after input, before post.
    '            Dim SqlCheck, StatusCode As String
    '            For Each dr As DataRow In dsItem.Tables(0).Rows
    '                If dr("Status").ToString = "CLID" Then
    '                    SqlCheck = String.Format("Select StatusCode from T_CLMaster with (nolock) where CLID='{0}' ", dr("CLID").ToString)
    '                    StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
    '                    If StatusCode = "0" Then
    '                        Return "Exists invalid CLID: " & dr("CLID").ToString
    '                    End If

    '                Else
    '                    SqlCheck = String.Format("Select TOP (1) StatusCode from T_CLMaster with (nolock) where BoxID='{0}' ", dr("BoxID").ToString)
    '                    StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
    '                    If StatusCode <> "1" Then
    '                        Return "Exists invalid BoxID: " & dr("BoxID").ToString
    '                    End If
    '                End If

    '                'Collect all DJs
    '                If dr("DJ").ToString <> "" Then
    '                    Dim DJdr() As DataRow
    '                    DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
    '                    If DJdr.Length = 0 Then
    '                        myDR = dtModel.NewRow
    '                        myDR("DJ") = dr("DJ").ToString
    '                        dtModel.Rows.Add(myDR)
    '                    End If
    '                End If
    '            Next

    '            Dim i As Integer

    '            Try
    '                Dim ds As DataSet = New DataSet()
    '                ds.Tables.Add("MatlList")
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("TranHeaderID", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_dn_num", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_move_order", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_mo_line_id", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_revision", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_number", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_quantity", System.Type.GetType("System.Decimal")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinv", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("o_success_flag", System.Type.GetType("System.String")))
    '                ds.Tables(0).Columns.Add(New Data.DataColumn("o_error_mssg", System.Type.GetType("System.String")))

    '                For i = 0 To dtPick.Rows.Count - 1
    '                    If dtPick.Rows(i)("PICKED_QTY") > 0 Then         'PICKEDQTY
    '                        myDR = ds.Tables(0).NewRow
    '                        myDR("TranHeaderID") = dtPick.Rows(i)("transaction_header_id").ToString
    '                        myDR("p_move_order") = dtPick.Rows(i)("MOVE_ORDER").ToString
    '                        myDR("p_mo_line_id") = dtPick.Rows(i)("MO_LINE_ID").ToString
    '                        myDR("p_item_num") = dtPick.Rows(i)("ITEM_NUMBER").ToString
    '                        myDR("p_subinv") = dtPick.Rows(i)("SOURCE_INV").ToString
    '                        If dtPick.Rows(i)("REVISION").ToString <> "" Then myDR("p_revision") = dtPick.Rows(i)("REVISION").ToString
    '                        If dtPick.Rows(i)("LOT_NUMBER").ToString <> "" Then myDR("p_lot_number") = dtPick.Rows(i)("LOT_NUMBER").ToString
    '                        If dtPick.Rows(i)("SOURCE_LOC").ToString <> "" Then myDR("p_locator") = dtPick.Rows(i)("SOURCE_LOC").ToString
    '                        myDR("p_dn_num") = ""

    '                        Dim Qty As Decimal = dtPick.Rows(i)("PICKED_QTY")
    '                        myDR("p_quantity") = Math.Round(Qty, 6)                                   'Keep 6 decimals for Qty and send to Oracle
    '                        ds.Tables(0).Rows.Add(myDR)
    '                    End If
    '                Next

    '                'Check if no data picked in the picking lists                   --   09/17/2014
    '                If ds.Tables(0).Rows.Count = 0 Then
    '                    MOrderPost = "Nothing Picked, please check your data."
    '                    ErrorLogging("ProdPicking-MOrderPost1", LoginData.User.ToUpper, Msg & MOrderPost, "I")
    '                    Return MOrderPost
    '                End If

    '                MOrderPost = ProcessMO(ds, LoginData)
    '                If MOrderPost <> "Y" Then Return MOrderPost
    '            Catch ex As Exception
    '                ErrorLogging("ProdPicking-MOrderPost2", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '                Return "Data update error, please contact IT."
    '            End Try


    '            Try

    '                'Read DJ Models from eTrace table or Oracle package
    '                If dtModel.Rows.Count > 0 Then
    '                    ReadDJModel(LoginData, dtModel)

    '                    For Each dr As DataRow In dsItem.Tables(0).Rows
    '                        If dr("DJ").ToString <> "" Then
    '                            Dim DJdr() As DataRow
    '                            DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
    '                            If DJdr.Length > 0 Then dr("Model") = DJdr(0)("Model").ToString
    '                        End If
    '                    Next
    '                    dsItem.Tables(0).AcceptChanges()
    '                End If

    '            Catch ex As Exception
    '                ErrorLogging("ProdPicking-MOrderPost4", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '                Return "Data update error, please contact IT."
    '            End Try


    '            Try
    '                Dim Sqlstr As String
    '                Sqlstr = String.Format("exec sp_MOrderPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, myMOData.MOType, LoginData.User.ToUpper, DStoXML(dsItem))
    '                MOrderPost = da.ExecuteScalar(Sqlstr).ToString
    '            Catch ex As Exception
    '                ErrorLogging("ProdPicking-sp_MOrderPost", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '                MOrderPost = "Data update error, please contact IT."
    '            End Try
    '            Return MOrderPost

    '        Catch ex As Exception
    '            ErrorLogging("ProdPicking-MOrderPost", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '            MOrderPost = "Data update error, please contact IT."
    '        End Try
    '    End Using

    'End Function
    Public Function MOrderPost(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            MOrderPost = ""

            Dim MO As String = myMOData.MO
            Dim dtPick As New DataTable
            Dim dsItem As DataSet = New DataSet
            Dim dsSubInv As DataSet = New DataSet

            Dim myDR As DataRow
            Dim dtModel As New DataTable
            dtModel.Columns.Add(New Data.DataColumn("DJ", System.Type.GetType("System.String")))
            dtModel.Columns.Add(New Data.DataColumn("Model", System.Type.GetType("System.String")))

            Dim dtMatl As New DataTable
            dtMatl.Columns.Add(New Data.DataColumn("DJ", System.Type.GetType("System.String")))
            dtMatl.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))

            Dim Msg As String = "MO: " & myMOData.MO & "; "

            Try

                dsItem.DataSetName = "Items"
                dsItem.Merge(myMOData.Items.Tables(1))

                dtPick = myMOData.Items.Tables(0).Copy
                myMOData.Items = New DataSet

                Dim TLAFlag As String = ""
                Dim ConfigRst As String = GetConfigValue("PIK013")
                If ConfigRst = "YES" Then
                    Dim SqlStr As String
                    SqlStr = String.Format("Select Description as SubInv from T_SysLOV with (nolock) where Name='BF Count' ")
                    dsSubInv = da.ExecuteDataSet(SqlStr, "SubInv")
                End If


                'Check if exists CLID's Status is 0, prevent user from slitting CLID after input, before post.
                Dim SqlCheck, StatusCode As String
                For Each dr As DataRow In dsItem.Tables(0).Rows
                    If dr("Status").ToString = "CLID" Then
                        SqlCheck = String.Format("Select StatusCode from T_CLMaster with (nolock) where CLID='{0}' ", dr("CLID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode = "0" Then
                            Return "Exists invalid CLID: " & dr("CLID").ToString
                        End If

                    Else
                        SqlCheck = String.Format("Select TOP (1) StatusCode from T_CLMaster with (nolock) where BoxID='{0}' ", dr("BoxID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode <> "1" Then
                            Return "Exists invalid BoxID: " & dr("BoxID").ToString
                        End If
                    End If

                    'Collect all DJs
                    If dr("DJ").ToString <> "" Then
                        Dim DJdr() As DataRow
                        DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
                        If DJdr.Length = 0 Then
                            myDR = dtModel.NewRow
                            myDR("DJ") = dr("DJ").ToString
                            dtModel.Rows.Add(myDR)
                        End If
                    End If

                    'Check if LTA DJ need to print Process Label or not for HH MO Confirm(ZS=YES / others =NO)
                    If ConfigRst = "YES" Then
                        Dim DestInv As String = dr("ToSubInv").ToString
                        If dsSubInv Is Nothing OrElse dsSubInv.Tables.Count = 0 OrElse dsSubInv.Tables(0).Rows.Count = 0 Then
                        ElseIf dsSubInv.Tables(0).Rows.Count > 0 AndAlso TLAFlag = "" Then
                            Dim Invdr() As DataRow
                            Invdr = dsSubInv.Tables(0).Select("SubInv = '" & DestInv & "'")
                            If Invdr.Length > 0 Then TLAFlag = "N"
                        End If
                        If TLAFlag = "" And DestInv.Contains("0 B") = True Then
                            TLAFlag = "Y"
                        End If
                    End If

                Next

                Dim i As Integer

                Try
                    Dim ds As DataSet = New DataSet()
                    ds.Tables.Add("MatlList")
                    ds.Tables(0).Columns.Add(New Data.DataColumn("TranHeaderID", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_dn_num", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_move_order", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_mo_line_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_revision", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_number", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_quantity", System.Type.GetType("System.Decimal")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinv", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_success_flag", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_error_mssg", System.Type.GetType("System.String")))

                    For i = 0 To dtPick.Rows.Count - 1
                        If dtPick.Rows(i)("PICKED_QTY") > 0 Then         'PICKEDQTY
                            myDR = ds.Tables(0).NewRow
                            myDR("TranHeaderID") = dtPick.Rows(i)("transaction_header_id").ToString
                            myDR("p_move_order") = dtPick.Rows(i)("MOVE_ORDER").ToString
                            myDR("p_mo_line_id") = dtPick.Rows(i)("MO_LINE_ID").ToString
                            myDR("p_item_num") = dtPick.Rows(i)("ITEM_NUMBER").ToString
                            myDR("p_subinv") = dtPick.Rows(i)("SOURCE_INV").ToString
                            If dtPick.Rows(i)("REVISION").ToString <> "" Then myDR("p_revision") = dtPick.Rows(i)("REVISION").ToString
                            If dtPick.Rows(i)("LOT_NUMBER").ToString <> "" Then myDR("p_lot_number") = dtPick.Rows(i)("LOT_NUMBER").ToString
                            If dtPick.Rows(i)("SOURCE_LOC").ToString <> "" Then myDR("p_locator") = dtPick.Rows(i)("SOURCE_LOC").ToString
                            myDR("p_dn_num") = ""

                            Dim Qty As Decimal = dtPick.Rows(i)("PICKED_QTY")
                            myDR("p_quantity") = Math.Round(Qty, 6)                                   'Keep 6 decimals for Qty and send to Oracle
                            ds.Tables(0).Rows.Add(myDR)
                        End If
                    Next

                    'Check if no data picked in the picking lists                   --   09/17/2014
                    If ds.Tables(0).Rows.Count = 0 Then
                        MOrderPost = "Nothing Picked, please check your data."
                        ErrorLogging("ProdPicking-MOrderPost1", LoginData.User.ToUpper, Msg & MOrderPost, "I")
                        Return MOrderPost
                    End If

                    MOrderPost = ProcessMO(ds, LoginData)
                    If MOrderPost <> "Y" Then Return MOrderPost
                Catch ex As Exception
                    ErrorLogging("ProdPicking-MOrderPost2", LoginData.User, Msg & ex.Message & ex.Source, "E")
                    Return "Data update error, please contact IT."
                End Try


                Try
                    'Read DJ Models from eTrace table or Oracle package
                    If dtModel.Rows.Count > 0 Then
                        ReadDJModel(LoginData, dtModel)

                        For Each dr As DataRow In dsItem.Tables(0).Rows
                            If dr("DJ").ToString <> "" Then
                                Dim DJdr() As DataRow
                                DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
                                If DJdr.Length > 0 Then dr("Model") = DJdr(0)("Model").ToString

                                'Collect PN for Process Label printing
                                If TLAFlag = "Y" Then
                                    Dim PNdr() As DataRow
                                    PNdr = dtMatl.Select("DJ = '" & dr("DJ").ToString & "' and MaterialNo ='" & dr("MaterialNo").ToString & "'")
                                    If PNdr.Length = 0 Then
                                        myDR = dtMatl.NewRow
                                        myDR("DJ") = dr("DJ").ToString
                                        myDR("MaterialNo") = dr("MaterialNo").ToString
                                        dtMatl.Rows.Add(myDR)
                                    End If
                                End If

                            End If
                        Next
                        dsItem.Tables(0).AcceptChanges()

                        'Check if this DJ is TLA DJ, if Yes, need to read data from Oracle and Print Process Label   -- 10/27/2017
                        If TLAFlag = "Y" Then
                            myMOData.Items.Tables.Add(dtModel)
                            myMOData.Items.Tables.Add(dtMatl)
                            DJReadProcessLabel(LoginData, myMOData)

                            'If BOM expand successfully, then Save Process Label Content to table T_RTSlip
                            If myMOData.Items Is Nothing OrElse myMOData.Items.Tables.Count = 0 Then
                            ElseIf myMOData.Items.Tables.Count > 0 Then
                                Dim dtCLID As New DataTable
                                dtCLID = dsItem.Tables(0).Copy
                                myMOData.Items.Tables.Add(dtCLID)

                                SaveProcessLabel(LoginData, myMOData)
                            End If
                        End If
                    End If

                Catch ex As Exception
                    ErrorLogging("ProdPicking-MOrderPost4", LoginData.User, Msg & ex.Message & ex.Source, "E")
                    Return "Data update error, please contact IT."
                End Try


                Try
                    Dim Sqlstr As String
                    Sqlstr = String.Format("exec sp_MOrderPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, myMOData.MOType, LoginData.User.ToUpper, DStoXML(dsItem))
                    MOrderPost = da.ExecuteScalar(Sqlstr).ToString
                Catch ex As Exception
                    ErrorLogging("ProdPicking-sp_MOrderPost", LoginData.User, Msg & ex.Message & ex.Source, "E")
                    MOrderPost = "Data update error, please contact IT."
                End Try

                'Light Off Slot if there has for PCB   -- 3/16/2018
                If myMOData.PickFlag = "Y" Then

                    Dim dsSlot As New DataSet("DS")
                    Dim dtSlot As DataTable
                    dtSlot = New DataTable("dtSlot")
                    dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                    dsSlot.Tables.Add(dtSlot)

                    Dim DR1() As DataRow
                    DR1 = dsItem.Tables(0).Select("Slot <> '' ")
                    If DR1.Length > 0 Then
                        For i = 0 To DR1.Length - 1
                            Dim Slotdr() As DataRow
                            Slotdr = dtSlot.Select("Slot ='" & DR1(i)("Slot").ToString & "'")
                            If Slotdr.Length > 0 Then Continue For

                            myDR = dtSlot.NewRow()
                            myDR("slot") = DR1(i)("Slot").ToString
                            dtSlot.Rows.Add(myDR)
                        Next

                        '==Code--0/1/2(Off/On/blink. data type: integer)
                        '==Interval--time(data type: integer. 0:not limit)
                        Dim Code As Integer = 0
                        Dim myWMS As WMS = New WMS
                        myWMS.LEDControlBySlot(dsSlot, Code, 0)
                    End If
                End If


                Return MOrderPost

            Catch ex As Exception
                ErrorLogging("ProdPicking-MOrderPost", LoginData.User, Msg & ex.Message & ex.Source, "E")
                MOrderPost = "Data update error, please contact IT."
            End Try
        End Using

    End Function

    Public Function MOSpecialPick(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim Items As DataSet = New DataSet
            Dim Msg As String = "Cancel allocation error for MO: " & myMOData.MO & " with item " & myMOData.SpecialItem & "; "

            Try

                Dim PickStatus As String = ""

                If Not myMOData.Items Is Nothing Then

                    Try
                        PickStatus = MOrderPost(LoginData, myMOData)
                    Catch ex As Exception
                        ErrorLogging("MOSpecialPick-MOrderPost", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        PickStatus = "Submit items error"
                    End Try
                    If PickStatus <> "Y" Then Return PickStatus
                End If


                'Call package to cancel the allocation for the special item, and extract the available Locator / Lot / Qty
                myMOData.Items = Nothing

                Try
                    Items = GetSMTOnHand(LoginData, myMOData)
                Catch ex As Exception
                    ErrorLogging("MOSpecialPick-GetSMTOnHand", LoginData.User, Msg & ex.Message & ex.Source, "E")
                    Return Msg
                End Try
                If Items Is Nothing OrElse Items.Tables.Count = 0 Then Return Msg

                'Return error message if only one column found from package
                If Items.Tables(0).Columns.Count = 1 Then
                    PickStatus = Items.Tables(0).Rows(0)(0).ToString
                    ErrorLogging("ProdPicking-MOSpecialPick1", LoginData.User, Msg & PickStatus, "I")

                    If PickStatus.ToUpper.Contains("ORA-00054") Then
                        PickStatus = "One user opened Transact Move Order in Oracle for MO: " & myMOData.MO & " with item " & myMOData.SpecialItem
                    End If
                    Return PickStatus
                End If

                'Return error message if no available Qty found from package
                If Items.Tables(0).Rows.Count = 0 Then
                    PickStatus = "No Available Qty for Item " & myMOData.SpecialItem & " in Source SubInv " & myMOData.SpecialSubInv & " for MO " & myMOData.MO
                    ErrorLogging("ProdPicking-MOSpecialPick2", LoginData.User, PickStatus, "I")
                    Return PickStatus
                End If

                myMOData.Items = Items
                MOSpecialPick = "Y"

            Catch ex As Exception
                ErrorLogging("ProdPicking-MOSpecialPick", LoginData.User, ex.Message & ex.Source, "E")
                Return Msg
            End Try

        End Using

    End Function

    Public Function MOSpecialPost(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            MOSpecialPost = ""

            Dim dtPick As New DataTable
            Dim dsItem As DataSet = New DataSet

            Dim myDR As DataRow
            Dim dtModel As New DataTable
            dtModel.Columns.Add(New Data.DataColumn("DJ", System.Type.GetType("System.String")))
            dtModel.Columns.Add(New Data.DataColumn("Model", System.Type.GetType("System.String")))

            Dim Msg As String = "MO: " & myMOData.MO & " with item " & myMOData.SpecialItem & "; "

            Try

                dsItem.DataSetName = "Items"
                dsItem.Merge(myMOData.Items.Tables(1))

                dtPick = myMOData.Items.Tables(0).Copy

                'Check if exists CLID's Status is 0, prevent user from slitting CLID after input, before post.
                Dim SqlCheck, StatusCode As String
                For Each dr As DataRow In dsItem.Tables(0).Rows
                    If dr("Status").ToString = "CLID" Then
                        SqlCheck = String.Format("Select StatusCode from T_CLMaster with (nolock) where CLID='{0}' ", dr("CLID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode = "0" Then
                            Return "Exists invalid CLID: " & dr("CLID").ToString
                        End If

                    Else
                        SqlCheck = String.Format("Select TOP (1) StatusCode from T_CLMaster with (nolock) where BoxID='{0}' ", dr("BoxID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode <> "1" Then
                            Return "Exists invalid BoxID: " & dr("BoxID").ToString
                        End If
                    End If

                    'Collect all DJs
                    If dr("DJ").ToString <> "" Then
                        Dim DJdr() As DataRow
                        DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
                        If DJdr.Length = 0 Then
                            myDR = dtModel.NewRow
                            myDR("DJ") = dr("DJ").ToString
                            dtModel.Rows.Add(myDR)
                        End If
                    End If
                Next


                Dim PickStatus As String = ""
                Dim ErrMsg As String = "Create allocation error for " & Msg

                Try
                    Dim i As Integer
                    For i = 0 To dtPick.Rows.Count - 1
                        If dtPick.Rows(i).RowState = DataRowState.Unchanged Then
                            dtPick.Rows(i).SetAdded()
                        End If
                    Next

                    myMOData.Items = New DataSet
                    myMOData.Items.Tables.Add(dtPick)

                    'Create MO Allocation for the specified Item
                    PickStatus = CreateMOAllocation(LoginData, myMOData)
                    If PickStatus <> "Y" Then Return PickStatus
                Catch ex As Exception
                    ErrorLogging("MOSpecialPost-CreateMOAllocation", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try


                PickStatus = ""
                ErrMsg = "Confirm Allocation error for " & Msg

                'Try
                '    myMOData.Items = Nothing

                '    'Confirm MO Allocation for the specified Item
                '    PickStatus = PostMOAllocation(LoginData, myMOData)
                '    If PickStatus <> "Y" Then Return PickStatus
                'Catch ex As Exception
                '    ErrorLogging("MOSpecialPost-PostMOAllocation", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                '    Return ErrMsg
                'End Try

                Try
                    Dim dtMO As New DataTable               'MatlList
                    dtMO.Columns.Add(New Data.DataColumn("MO", System.Type.GetType("System.String")))
                    dtMO.Columns.Add(New Data.DataColumn("Item", System.Type.GetType("System.String")))
                    dtMO.Columns.Add(New Data.DataColumn("Qty", System.Type.GetType("System.Decimal")))
                    dtMO.Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
                    dtMO.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))

                    Dim i As Integer
                    Dim DR() As DataRow

                    For i = 0 To dtPick.Rows.Count - 1
                        DR = dtMO.Select("Item = '" & dtPick.Rows(i)("item_num") & "'")
                        If DR.Length = 0 Then
                            myDR = dtMO.NewRow
                            myDR("MO") = myMOData.MO
                            myDR("Item") = dtPick.Rows(i)("item_num")
                            myDR("Status") = ""
                            myDR("Message") = ""

                            Dim Qty As Decimal = dtPick.Rows(i)("picked_qty")
                            myDR("Qty") = Math.Round(Qty, 6)
                            dtMO.Rows.Add(myDR)
                        Else
                            Dim TotalQty As Decimal = DR(0)("Qty") + dtPick.Rows(i)("picked_qty")
                            DR(0)("Qty") = Math.Round(TotalQty, 6)
                            DR(0).AcceptChanges()
                            DR(0).SetAdded()
                        End If
                    Next

                    myMOData.Items = New DataSet
                    myMOData.Items.Tables.Add(dtMO)

                    'Confirm MO Allocation for the specified Item
                    PickStatus = PostSMTMOAllocation(LoginData, myMOData)
                    If PickStatus <> "Y" Then Return PickStatus
                Catch ex As Exception
                    ErrorLogging("MOSpecialPost-PostSMTMOAllocation", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try


                ErrMsg = "Data update error for " & Msg

                Try
                    'Read DJ Models from eTrace table or Oracle package
                    If dtModel.Rows.Count > 0 Then
                        ReadDJModel(LoginData, dtModel)

                        For Each dr As DataRow In dsItem.Tables(0).Rows
                            If dr("DJ").ToString <> "" Then
                                Dim DJdr() As DataRow
                                DJdr = dtModel.Select("DJ = '" & dr("DJ").ToString & "'")
                                If DJdr.Length > 0 Then dr("Model") = DJdr(0)("Model").ToString
                            End If
                        Next
                        dsItem.Tables(0).AcceptChanges()
                    End If

                Catch ex As Exception
                    ErrorLogging("ProdPicking-MOSpecialPost1", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try

                Try

                    Dim Sqlstr As String
                    Sqlstr = String.Format("exec sp_MOrderPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, myMOData.MOType, LoginData.User.ToUpper, DStoXML(dsItem))
                    MOSpecialPost = da.ExecuteScalar(Sqlstr).ToString

                Catch ex As Exception
                    ErrorLogging("MOSpecialPost-sp_MOrderPost", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try
                Return MOSpecialPost

            Catch ex As Exception
                MOSpecialPost = "Data update error for " & Msg
                ErrorLogging("ProdPicking-MOSpecialPost", LoginData.User, MOSpecialPost & ex.Message & ex.Source, "E")
            End Try

        End Using

    End Function

    Public Function GetSMTOnHand(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim Msg As String = "MO: " & myMOData.MO & " with item " & myMOData.SpecialItem & "; "



            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                Dim ConfigID As String = "MOC002"
                Dim GetExpDFlag As String = ""
                GetExpDFlag = GetConfigValue(ConfigID)

                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("MatlList")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.XXETR_wip_pkg.get_smt_onhand_info"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_mo_num", OracleType.VarChar, 50)).Value = myMOData.MO
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_item", OracleType.VarChar, 50)).Value = myMOData.SpecialItem
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_subinv", OracleType.VarChar, 50)).Value = myMOData.SpecialSubInv

                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_lot_expiration_flag", OracleType.VarChar, 50)).Value = GetExpDFlag   '"N"

                oda.SelectCommand.Parameters.Add("o_cur_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "MatlList")
                oda.SelectCommand.Connection.Close()

                Return ds

            Catch ex As Exception
                ErrorLogging("ProdPicking-GetSMTOnHand", LoginData.User.ToUpper, Msg & ex.Message & ex.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function CreateMOAllocation(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Insert()
            Dim Msg As String = "Create allocation error for MO: " & myMOData.MO & " with item " & myMOData.SpecialItem

            Dim ds As DataSet = New DataSet()

            Try
                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                Dim ConfigID As String = "MOC002"
                Dim GetExpDFlag As String = ""
                GetExpDFlag = GetConfigValue(ConfigID)

                ds = myMOData.Items
                ds.Tables(0).Columns.Add(New Data.DataColumn("Flag", System.Type.GetType("System.String")))
                ds.Tables(0).Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))

                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.XXETR_wip_pkg.create_smt_allocation"
                oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32).Value = myMOData.OrgID
                oda.InsertCommand.Parameters.Add("p_item_id", OracleType.Int32).Value = myMOData.ItemID
                oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32).Value = LoginData.UserID
                oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv

                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                oda.InsertCommand.Parameters.Add("p_lot_expiration_flag", OracleType.VarChar, 50).Value = GetExpDFlag   '"N"

                oda.InsertCommand.Parameters.Add("p_moline_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_subinv", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_locator_id", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_lot_num", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                oda.InsertCommand.Parameters("p_moline_id").SourceColumn = "mo_line_id"
                oda.InsertCommand.Parameters("p_subinv").SourceColumn = "subinv"
                oda.InsertCommand.Parameters("p_locator_id").SourceColumn = "locator_id"
                oda.InsertCommand.Parameters("p_lot_num").SourceColumn = "lot_num"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "picked_qty"
                oda.InsertCommand.Parameters("o_success_flag").SourceColumn = "Flag"
                oda.InsertCommand.Parameters("o_error_mssg").SourceColumn = "Message"

                oda.InsertCommand.Connection.Open()
                oda.Update(ds.Tables(0))
                oda.InsertCommand.Connection.Close()

                Dim DR() As DataRow = Nothing
                DR = ds.Tables(0).Select(" Flag = 'N' or Flag = ' ' or Flag IS Null ")
                If DR.Length = 0 Then Return "Y"

                'Record and Return Error message
                Dim i As Integer
                Dim ErrMsg As String = Msg
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString <> "" Then
                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString
                    End If
                Next
                ErrorLogging("ProdPicking-CreateMOAllocation1", LoginData.User.ToUpper, ErrMsg, "I")
                Return ErrMsg

            Catch ex As Exception
                ErrorLogging("ProdPicking-CreateMOAllocation", LoginData.User.ToUpper, Msg & "; " & ex.Message & ex.Source, "E")
                Return Msg
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function PostMOAllocation(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim ErrMsg As String = "Confirm Allocation error for MO: " & myMOData.MO & " with item " & myMOData.SpecialItem

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.process_smt_mo"
                OC.Parameters.Add("p_org_id", OracleType.VarChar).Value = myMOData.OrgID
                OC.Parameters.Add("p_moheader_id", OracleType.VarChar, 20).Value = myMOData.MOHeaderID
                OC.Parameters.Add("p_item_id", OracleType.VarChar, 20).Value = myMOData.ItemID
                OC.Parameters.Add("o_success_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Dim Flag, Msg As String
                Flag = OC.Parameters("o_success_flag").Value.ToString
                Msg = OC.Parameters("o_error_mssg").Value.ToString
                OC.Connection.Close()

                If Flag = "Y" Then        ' Y: Confirm Allocation successful; N: Confirm Allocation failed
                    Return Flag
                Else
                    ErrorLogging("ProdPicking-PostMOAllocation1", LoginData.User.ToUpper, ErrMsg & " and error message; " & Msg, "I")
                    Return Msg
                End If

            Catch ex As Exception
                ErrorLogging("ProdPicking-PostMOAllocation", LoginData.User.ToUpper, ErrMsg & "; " & ex.Message & ex.Source, "E")
                Return ErrMsg
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try

        End Using

    End Function

    Public Function PostSMTMOAllocation(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim ds As DataSet = New DataSet()
            ds = myMOData.Items

            Try
                Dim MO As String = ds.Tables(0).Rows(0)("MO").ToString

                Dim oda As OracleDataAdapter
                Dim comm As OracleCommand

                Dim OrgID As String = GetOrgID(LoginData.OrgCode)
                oda = New OracleDataAdapter
                comm = da.Ora_Command_Trans()

                Try
                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.XXETR_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                    comm.ExecuteOracleNonQuery("")
                    comm.Parameters.Clear()

                    comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                    comm.Parameters.Add("p_org_id", OracleType.VarChar, 20).Value = OrgID                        'LoginData.OrgCode
                    comm.Parameters.Add("p_move_order", OracleType.VarChar, 50).Value = MO
                    comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_quantity", OracleType.Double)
                    comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
                    comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

                    comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                    comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                    comm.Parameters("p_item_num").SourceColumn = "Item"
                    comm.Parameters("p_quantity").SourceColumn = "Qty"

                    comm.Parameters("o_success_flag").SourceColumn = "Status"
                    comm.Parameters("o_error_mssg").SourceColumn = "Message"

                    oda.InsertCommand = comm
                    oda.Update(ds.Tables(0))

                    If ds.Tables(0).Rows.Count = 0 Then
                        comm.Transaction.Rollback()
                        comm.Connection.Close()
                        PostSMTMOAllocation = "No data is submitted in Oracle"
                        Return PostSMTMOAllocation
                    End If

                    Dim DR() As DataRow = Nothing
                    DR = ds.Tables(0).Select(" Status = 'N' or Status = ' ' or Status IS Null ")
                    If DR.Length = 0 Then
                        comm.Transaction.Commit()
                        comm.Connection.Close()
                        PostSMTMOAllocation = "Y"                                'Oracle Data update successfully and set Flag = "Y"
                        Return PostSMTMOAllocation
                    End If

                    'Rollback and record error message if Oracle Data update failed
                    comm.Transaction.Rollback()
                    comm.Connection.Close()

                    Dim i As Integer
                    For i = 0 To DR.Length - 1
                        Dim ErrMsg As String = ""
                        ErrMsg = "MO " & MO & " for Item: " & DR(i)("Item").ToString & " with IssuedQty " & DR(i)("Qty").ToString & " and flag: " & DR(i)("Status").ToString & " and error message; " & DR(i)("Message").ToString
                        ErrorLogging("ProdPicking-PostSMTMOAllocation1", LoginData.User, ErrMsg, "I")
                    Next

                    PostSMTMOAllocation = DR(0)("Message").ToString
                    If PostSMTMOAllocation = "" Then
                        PostSMTMOAllocation = "No data is submitted in Oracle"
                        Return PostSMTMOAllocation
                    End If

                    If PostSMTMOAllocation.ToUpper.Contains("ORA-00054") Then
                        PostSMTMOAllocation = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(0)("Item").ToString
                    Else
                        PostSMTMOAllocation = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
                    End If
                    Return PostSMTMOAllocation                                'Return Error message

                Catch ex As Exception
                    ErrorLogging("ProdPicking-PostSMTMOAllocation", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                    comm.Transaction.Rollback()
                    PostSMTMOAllocation = ex.Message
                Finally
                    If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
                End Try

            Catch ex As Exception
                ErrorLogging("PostSMTMOAllocation-Exception", LoginData.User, ex.Message & ex.Source, "E")
                PostSMTMOAllocation = ex.Message
            End Try

        End Using

    End Function

    Public Function MOEarlyLotPost(ByVal LoginData As ERPLogin, ByVal myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            MOEarlyLotPost = ""

            Dim dtPick As New DataTable
            Dim dsItem As DataSet = New DataSet

            Dim myDR As DataRow
            Dim dtModel As New DataTable
            dtModel.Columns.Add(New Data.DataColumn("DJ", System.Type.GetType("System.String")))
            dtModel.Columns.Add(New Data.DataColumn("Model", System.Type.GetType("System.String")))

            Dim Msg As String = "MO: " & myMOData.MO & " with item " & myMOData.SpecialItem & "; "

            Try

                dsItem.DataSetName = "Items"
                dsItem.Merge(myMOData.Items.Tables(1))

                dtPick = myMOData.Items.Tables(0).Copy

                'Check if exists CLID's Status is 0, prevent user from slitting CLID after input, before post.
                Dim SqlCheck, StatusCode As String
                For Each dr As DataRow In dsItem.Tables(0).Rows
                    If dr("Status").ToString = "CLID" Then
                        SqlCheck = String.Format("Select StatusCode from T_CLMaster with (nolock) where CLID='{0}' ", dr("CLID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode = "0" Then
                            Return "Exists invalid CLID: " & dr("CLID").ToString
                        End If

                    Else
                        SqlCheck = String.Format("Select TOP (1) StatusCode from T_CLMaster with (nolock) where BoxID='{0}' ", dr("BoxID").ToString)
                        StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                        If StatusCode <> "1" Then
                            Return "Exists invalid BoxID: " & dr("BoxID").ToString
                        End If
                    End If
                Next


                Dim i As Integer
                Dim PickStatus As String = ""
                Dim ErrMsg As String = "Change Lot allocation/confirm error for " & Msg

                Dim DJModel As String = ""
                Dim DJNo As String = myMOData.Items.Tables(0).Rows(0)("DJ_Name").ToString

                Try
                    For i = 0 To dtPick.Rows.Count - 1
                        If dtPick.Rows(i).RowState = DataRowState.Unchanged Then
                            dtPick.Rows(i).SetAdded()
                        End If
                    Next

                    myMOData.Items = New DataSet
                    myMOData.Items.Tables.Add(dtPick)

                    'Change MO Lot Allocation and confirm for the specified Item
                    PickStatus = ChangeMOLotAllocation(LoginData, myMOData)
                    If PickStatus <> "Y" Then Return PickStatus

                    DJModel = myMOData.Items.Tables(0).Rows(0)("DJModel").ToString.Trim
                Catch ex As Exception
                    ErrorLogging("MOEarlyLotPost-ChangeMOLotAllocation", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try

                'Get DJ/Model from Oracle pkg
                If DJModel.Contains("/") AndAlso DJModel.Length > 1 Then
                    Dim DJArry() As String = Split(DJModel, "/")
                    DJNo = DJArry(0).ToString
                    DJModel = DJArry(1).ToString
                Else
                    If DJNo <> "" Then
                        myDR = dtModel.NewRow
                        myDR("DJ") = DJNo
                        dtModel.Rows.Add(myDR)

                        'Read DJ Models from eTrace table or Oracle package
                        ReadDJModel(LoginData, dtModel)
                        DJModel = dtModel.Rows(0)("Model").ToString
                    End If
                End If

                For i = 0 To dsItem.Tables(0).Rows.Count - 1
                    dsItem.Tables(0).Rows(i)("DJ") = DJNo
                    dsItem.Tables(0).Rows(i)("Model") = DJModel
                Next
                dsItem.Tables(0).AcceptChanges()


                ErrMsg = "Data update error for " & Msg

                Try

                    Dim Sqlstr As String
                    Sqlstr = String.Format("exec sp_MOrderPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, myMOData.MOType, LoginData.User.ToUpper, DStoXML(dsItem))
                    MOEarlyLotPost = da.ExecuteScalar(Sqlstr).ToString

                Catch ex As Exception
                    ErrorLogging("MOEarlyLotPost-sp_MOrderPost", LoginData.User, ErrMsg & ex.Message & ex.Source, "E")
                    Return ErrMsg
                End Try
                Return MOEarlyLotPost

            Catch ex As Exception
                MOEarlyLotPost = "Data update error for " & Msg
                ErrorLogging("ProdPicking-MOEarlyLotPost", LoginData.User, MOEarlyLotPost & ex.Message & ex.Source, "E")
            End Try

        End Using

    End Function

    Public Function ChangeMOLotAllocation(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Insert()
            Dim Msg As String = "Change Lot allocation/confirm error for MO: " & myMOData.MO & " with item " & myMOData.SpecialItem

            Dim ds As DataSet = New DataSet()

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                Dim ConfigID As String = "MOC002"
                Dim GetExpDFlag As String = ""
                GetExpDFlag = GetConfigValue(ConfigID)

                ds = myMOData.Items
                ds.Tables(0).Columns.Add(New Data.DataColumn("Flag", System.Type.GetType("System.String")))
                ds.Tables(0).Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
                ds.Tables(0).Columns.Add(New Data.DataColumn("DJModel", System.Type.GetType("System.String")))

                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.XXETR_wip_pkg.change_smtmo_lotnum"
                oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32).Value = OrgID    'myMOData.OrgID
                oda.InsertCommand.Parameters.Add("p_moheader_id", OracleType.Int32).Value = myMOData.MOHeaderID
                oda.InsertCommand.Parameters.Add("p_item_id", OracleType.Int32).Value = myMOData.ItemID
                oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32).Value = LoginData.UserID
                oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv

                'Consider Oracle Lot ExpDate or not while getting Available Qty from Oracle, default value = Y
                oda.InsertCommand.Parameters.Add("p_lot_expiration_flag", OracleType.VarChar, 50).Value = GetExpDFlag    '"N"

                'oda.InsertCommand.Parameters.Add("p_moline_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_subinv", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_lot_to", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_djmodel", OracleType.VarChar, 100).Direction = ParameterDirection.Output

                'oda.InsertCommand.Parameters("p_moline_id").SourceColumn = "mo_line_id"
                oda.InsertCommand.Parameters("p_subinv").SourceColumn = "SOURCE_INV"
                oda.InsertCommand.Parameters("p_locator").SourceColumn = "SOURCE_LOC"
                oda.InsertCommand.Parameters("p_lot_to").SourceColumn = "LOT_NUMBER"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "PICKED_QTY"
                oda.InsertCommand.Parameters("o_success_flag").SourceColumn = "Flag"
                oda.InsertCommand.Parameters("o_error_mssg").SourceColumn = "Message"
                oda.InsertCommand.Parameters("o_djmodel").SourceColumn = "DJModel"

                oda.InsertCommand.Connection.Open()
                oda.Update(ds.Tables(0))
                oda.InsertCommand.Connection.Close()

                Dim DR() As DataRow = Nothing
                DR = ds.Tables(0).Select(" Flag = 'N' or Flag = ' ' or Flag IS Null ")
                If DR.Length = 0 Then Return "Y"

                'Record and Return Error message
                Dim i As Integer
                Dim ErrMsg As String = Msg
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString <> "" Then
                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString
                    End If
                Next
                ErrorLogging("ProdPicking-ChangeMOLotAllocation1", LoginData.User.ToUpper, ErrMsg, "I")
                Return ErrMsg

            Catch ex As Exception
                ErrorLogging("ProdPicking-ChangeMOLotAllocation", LoginData.User.ToUpper, Msg & "; " & ex.Message & ex.Source, "E")
                Return Msg
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

        End Using

    End Function


    Public Function ReadDJModel(ByVal LoginData As ERPLogin, ByVal dtModel As DataTable) As DataTable
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i As Integer
                Dim SqlStr As String
                Dim DJStr As String = ""

                For i = 0 To dtModel.Rows.Count - 1
                    Dim Model As String = ""
                    Dim DJ As String = dtModel.Rows(i)("DJ").ToString

                    'Read DJ Product from table T_PO_CLID first
                    SqlStr = String.Format("Select TOP (1) Product from T_PO_CLID with (nolock) where Product <>'' and OrgCode = '{0}' and PO = '{1}'", LoginData.OrgCode, DJ)
                    Model = Convert.ToString(da.ExecuteScalar(SqlStr))
                    dtModel.Rows(i)("Model") = Model
                    If Model <> "" Then Continue For

                    'Read DJ Product from table T_DJInfo if T_PO_CLID not found the first time
                    SqlStr = String.Format("Select TOP (1) Model from T_DJInfo with (nolock) where OrgCode = '{0}' and PO = '{1}'", LoginData.OrgCode, DJ)
                    Model = Convert.ToString(da.ExecuteScalar(SqlStr))
                    dtModel.Rows(i)("Model") = Model
                    If Model <> "" Then Continue For

                    If DJStr = "" Then
                        DJStr = DJ
                    Else
                        DJStr = DJStr & "," & DJ
                    End If

                    'Record errorlog if T_DJInfo not found
                    Dim ErrMsg As String
                    ErrMsg = "Could not find DJ Model from table T_DJInfo for DJ " & DJ
                    ErrorLogging("ProdPicking-ReadDJModel", LoginData.User, ErrMsg, "I")

                Next

                'Read DJ Module from Oracle according to SP: sp_DJmodel_info
                If DJStr <> "" Then
                    Dim ds As New DataSet()
                    Try
                        SqlStr = String.Format("exec dbo.sp_DJmodel_info '{0}','{1}'", LoginData.OrgCode, DJStr)
                        ds = da.ExecuteDataSet(SqlStr, "DJModel")
                    Catch ex As Exception
                        ErrorLogging("ProdPicking-Call:sp_DJmodel_info", LoginData.User, ex.Message & ex.Source, "E")
                        ds = Nothing
                    End Try

                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    ElseIf ds.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dtModel.Rows.Count - 1
                            If dtModel.Rows(i)("Model") = "" Then
                                Dim DR() As DataRow
                                DR = ds.Tables(0).Select("DJ = '" & dtModel.Rows(i)("DJ") & "'")
                                If DR.Length > 0 Then
                                    dtModel.Rows(i)("Model") = DR(0)("Model").ToString
                                End If
                            End If
                        Next
                    End If
                End If

            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadDJModel", LoginData.User, ex.Message & ex.Source, "E")
            End Try

            Return dtModel

        End Using

    End Function

    Public Function ReadPOrderList(ByVal ProdFloor As String, ByVal Open As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim ds As DataSet = New DataSet()
                Dim Sqlstr As String
                If Open = False Then    'Show Open Pick Orders
                    Sqlstr = String.Format("Select A.PDTO, A.ProdLine, A.CreatedOn AS CreatedDate, A.Status, A.SupplyType, A.DestSubInv, A.DestLocator, A.ReasonCode, B.PO, B.Product, B.BuildQty FROM T_PDTOHeader AS A with (nolock) INNER JOIN T_PDTO_PO AS B with (nolock) ON A.PDTO = B.PDTO where A.ProdFloor ='{0}' and OrgCode = '{1}' and A.Status='1' and datediff(month,CreatedOn,getdate())<=2 ORDER BY A.PDTO", ProdFloor, LoginData.OrgCode) '
                Else                    'Show all Pick Orders
                    Sqlstr = String.Format("Select A.PDTO, A.ProdLine, A.CreatedOn AS CreatedDate, A.Status, A.SupplyType, A.DestSubInv, A.DestLocator, A.ReasonCode, B.PO, B.Product, B.BuildQty FROM T_PDTOHeader AS A with (nolock) INNER JOIN T_PDTO_PO AS B with (nolock) ON A.PDTO = B.PDTO where A.ProdFloor ='{0}' and OrgCode = '{1}' and datediff(month,CreatedOn,getdate())<=2 ORDER BY A.PDTO", ProdFloor, LoginData.OrgCode) '
                End If
                ds = da.ExecuteDataSet(Sqlstr, "PDTOs")

                Return ds
            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadPOrderList", LoginData.User.ToUpper, "ProdFloor:" & ProdFloor & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function ReadPOrderItems(ByVal PickOrder As String, ByVal ShowAllMatls As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim ConfigRst As String = GetConfigValue("PIK012")
                Dim ds As DataSet = New DataSet()

                'First Lock the Pick Order ( Set Status = 2 in table T_PDTOHeader ) if user do picking according to T_Config ( =YES )  -- 5/23/2017
                If ConfigRst = "YES" Then
                    da.ExecuteNonQuery(String.Format("Update T_PDTOHeader set Status='2', ChangedOn=getDate(), ChangedBy='{0}' where PDTO='{1}'", LoginData.User.ToUpper, PickOrder))
                End If

                Dim Sqlstr As String = ""
                If ShowAllMatls = True Then
                    Sqlstr = String.Format("Select PDTO,Item,Material,MaterialRevision,ReqQty,PickedQty,Status,Manufacturer,MPN,CurrQty=0.000, CHFlag='', DJStatus='' from T_PDTOItem with (nolock) where PDTO = '{0}' ORDER BY Material", PickOrder)
                Else
                    Sqlstr = String.Format("Select PDTO,Item,Material,MaterialRevision,ReqQty,PickedQty,Status,Manufacturer,MPN,CurrQty=0.000, CHFlag='', DJStatus='' from T_PDTOItem with (nolock) where PDTO = '{0}' and Status='1' ORDER BY Material", PickOrder)
                End If
                ds = da.ExecuteDataSet(Sqlstr, "PDTOItems")

                Dim CHFlag As String = GetConfigValue("CLID014")
                If CHFlag = "YES" Then
                    Dim DJNo As String = ""
                    Dim DJStatus As String = ""

                    Sqlstr = String.Format("Select PO from T_PDTO_PO with (nolock) where PDTO = '{0}' ", PickOrder)
                    DJNo = Convert.ToString(da.ExecuteScalar(Sqlstr))

                    'Read DJ Status from HuaWei Web Service 
                    DJStatus = CheckCHDJ(DJNo, "")                         ' DJStatus = "0"  means OK

                    ds.Tables(0).Rows(0)("CHFlag") = CHFlag
                    ds.Tables(0).Rows(0)("DJStatus") = DJStatus
                End If

                Return ds

            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadPOrderItems", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function UpdatePOrderHeader(ByVal PickOrder As String, ByVal Status As String, ByVal LoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                Dim ConfigRst As String = GetConfigValue("PIK012")
                If ConfigRst = "YES" Then
                    Dim sqlstr As String
                    sqlstr = String.Format("Update T_PDTOHeader set Status='{0}',ChangedOn=getDate(), ChangedBy='{1}' where PDTO='{2}' ", Status, LoginData.User.ToUpper, PickOrder)
                    da.ExecuteNonQuery(sqlstr)
                End If

                Return True
            Catch ex As Exception
                ErrorLogging("ProdPicking-UpdatePOrderHeader", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function PickOrderPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            PickOrderPost = ""

            Dim PickOrder, SupplyType As String
            PickOrder = Items.Tables(1).Rows(0)("PickOrder").ToString.ToUpper
            SupplyType = Items.Tables(1).Rows(0)("SupplyType").ToString.ToUpper

            Try
                Dim SqlCheck As String
                Dim dsSubInv As DataSet = New DataSet

                SqlCheck = String.Format("Select * from T_BackflushSubinv ")
                dsSubInv = da.ExecuteDataSet(SqlCheck, "BFSubInv")

                Dim i As Integer
                Dim myDR As DataRow
                Dim DR() As DataRow = Nothing
                Dim ds As DataSet = New DataSet()

                If SupplyType = "PUSH" Then                                         'WIP Component Issue in Oracle

                    ds.Tables.Add("MatlList")
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_dj_name", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_rev", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinventory", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_number", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_issue_quantity", System.Type.GetType("System.Decimal")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_uom", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_reason", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_success_flag", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_error_msgg", System.Type.GetType("System.String")))
                    ds.AcceptChanges()

                    For i = 0 To Items.Tables(1).Rows.Count - 1
                        DR = ds.Tables(0).Select("p_item_num = '" & Items.Tables(1).Rows(i)("MaterialNo").ToString & _
                                               "' and p_item_rev = '" & Items.Tables(1).Rows(i)("MaterialRevision").ToString & _
                                               "' and p_subinventory = '" & Items.Tables(1).Rows(i)("SubInv").ToString & _
                                               "' and p_locator = '" & Items.Tables(1).Rows(i)("Locator").ToString & _
                                               "' and p_lot_number = '" & Items.Tables(1).Rows(i)("RTLot").ToString & "'")
                        If DR.Length > 0 Then
                            'DR(0)("p_issue_quantity") = DR(0)("p_issue_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            Dim Qty As Decimal = DR(0)("p_issue_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            DR(0)("p_issue_quantity") = Math.Round(Qty, 5)          'Keep 5 decimals for Qty and send to Oracle
                        Else
                            myDR = ds.Tables(0).NewRow
                            myDR("p_dj_name") = Items.Tables(1).Rows(i)("DJ").ToString
                            myDR("p_item_num") = Items.Tables(1).Rows(i)("MaterialNo").ToString
                            myDR("p_item_rev") = Items.Tables(1).Rows(i)("MaterialRevision").ToString
                            myDR("p_subinventory") = Items.Tables(1).Rows(i)("SubInv").ToString
                            myDR("p_locator") = Items.Tables(1).Rows(i)("Locator").ToString
                            myDR("p_lot_number") = Items.Tables(1).Rows(i)("RTLot").ToString
                            myDR("p_issue_quantity") = Items.Tables(1).Rows(i)("CLIDQty")
                            myDR("p_uom") = Items.Tables(1).Rows(i)("BaseUOM").ToString
                            myDR("p_reason") = Items.Tables(1).Rows(i)("ReasonCode").ToString
                            myDR("o_success_flag") = ""
                            myDR("o_error_msgg") = ""
                            ds.Tables(0).Rows.Add(myDR)
                        End If
                    Next
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(i).SetAdded()
                        End If
                    Next

                    Dim dsPush As DataSet = New DataSet
                    Try
                        dsPush = OracleWIPIssue(ds, LoginData)
                    Catch ex As Exception
                        ErrorLogging("PickOrderPost-OracleWIPIssue", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                        Return ex.Message
                    End Try

                    DR = Nothing
                    DR = dsPush.Tables(0).Select("o_success_flag = 'N' or o_success_flag = ' ' or o_success_flag IS Null ")
                    If DR.Length > 0 Then
                        PickOrderPost = DR(0)("o_error_msgg").ToString
                        If PickOrderPost = "" Then PickOrderPost = "No data is submitted in Oracle"
                        Return PickOrderPost
                    End If

                ElseIf SupplyType = "PULL" OrElse SupplyType = "MSB" OrElse SupplyType = "PRE-WORK" Then           'SubInv Transfer in Oracle

                    Dim TransactionID As Long
                    TransactionID = CLng(GetTranHeaderID(LoginData, "HDID"))

                    ''Record the TranHeaderID to trace if it's duplicate or not -- 07/03/2013
                    'Dim Sqlstr As String
                    'Dim Remark As String = Items.Tables(1).Rows(0)("DJ").ToString & " / " & PickOrder
                    'Sqlstr = String.Format("INSERT INTO T_TranHeaderID (TranHeaderID,CreatedOn,CreatedBy,Remark) values ('{0}',getdate(),'{1}','{2}') ", TransactionID, LoginData.User.ToUpper, Remark)
                    'da.ExecuteNonQuery(Sqlstr)


                    ds.Tables.Add("transfer_table")
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_timeout", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_organization_code", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_transaction_header_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_transaction_uom", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_source_line_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_source_header_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_user_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_segment1", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_revision", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinventory_source", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator_source", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinventory_destination", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator_destination", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_number", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_expiration_date", System.Type.GetType("System.DateTime")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_transaction_quantity", System.Type.GetType("System.Decimal")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_primary_quantity", System.Type.GetType("System.Decimal")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_reason_code", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_transaction_reference", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_return_status", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_return_message", System.Type.GetType("System.String")))
                    ds.AcceptChanges()

                    For i = 0 To Items.Tables(1).Rows.Count - 1
                        Dim RTLot As String = Items.Tables(1).Rows(i)("RTLot").ToString

                        'For SupplyType="PULL", check if the Source SubInv is contained in table T_BackflushSubinv, if yes, then set RTLot as: 00000 and send to Oracle
                        'In this case, package will auto search the available RTLot with early ExpDate under the source SubInv/Locator to do SubInv Transfer
                        If SupplyType = "PULL" AndAlso RTLot <> "" Then
                            If dsSubInv Is Nothing OrElse dsSubInv.Tables.Count = 0 OrElse dsSubInv.Tables(0).Rows.Count = 0 Then
                            ElseIf dsSubInv.Tables(0).Rows.Count > 0 Then
                                Dim BFdr() As DataRow = Nothing
                                BFdr = dsSubInv.Tables(0).Select("Subinventory = '" & Items.Tables(1).Rows(i)("SubInv").ToString & "'")
                                If BFdr.Length > 0 Then RTLot = "00000"
                            End If
                        End If

                        DR = ds.Tables(0).Select("p_item_segment1 = '" & Items.Tables(1).Rows(i)("MaterialNo").ToString & _
                                               "' and p_item_revision = '" & Items.Tables(1).Rows(i)("MaterialRevision").ToString & _
                                               "' and p_subinventory_source = '" & Items.Tables(1).Rows(i)("SubInv").ToString & _
                                               "' and p_locator_source = '" & Items.Tables(1).Rows(i)("Locator").ToString & _
                                               "' and p_lot_number = '" & RTLot & "'")

                        If DR.Length > 0 Then
                            'DR(0)("p_transaction_quantity") = DR(0)("p_transaction_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            'DR(0)("p_primary_quantity") = DR(0)("p_primary_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            Dim Qty, TranQty As Decimal
                            Qty = DR(0)("p_primary_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            TranQty = DR(0)("p_transaction_quantity") + Items.Tables(1).Rows(i)("CLIDQty")
                            DR(0)("p_primary_quantity") = Math.Round(Qty, 5)                  'Keep 5 decimals for Qty and send to Oracle
                            DR(0)("p_transaction_quantity") = Math.Round(TranQty, 5)          'Keep 5 decimals for Qty and send to Oracle
                        Else

                            myDR = ds.Tables(0).NewRow()
                            myDR("p_timeout") = "1800000"
                            myDR("p_organization_code") = LoginData.OrgCode
                            myDR("p_transaction_header_id") = TransactionID
                            myDR("p_transaction_uom") = Items.Tables(1).Rows(i)("BaseUOM").ToString
                            myDR("p_source_line_id") = "99"
                            myDR("p_source_header_id") = "99"
                            myDR("p_user_id") = LoginData.UserID
                            myDR("p_item_segment1") = Items.Tables(1).Rows(i)("MaterialNo").ToString
                            myDR("p_item_revision") = Items.Tables(1).Rows(i)("MaterialRevision").ToString
                            myDR("p_subinventory_source") = Items.Tables(1).Rows(i)("SubInv").ToString
                            myDR("p_locator_source") = Items.Tables(1).Rows(i)("Locator").ToString
                            myDR("p_subinventory_destination") = Items.Tables(1).Rows(i)("DestSubInv").ToString
                            myDR("p_locator_destination") = Items.Tables(1).Rows(i)("DestLocator").ToString
                            myDR("p_lot_number") = RTLot                  'Items.Tables(1).Rows(i)("RTLot")
                            myDR("p_lot_expiration_date") = IIf(Items.Tables(1).Rows(i)("ExpDate") Is DBNull.Value, DBNull.Value, Items.Tables(1).Rows(i)("ExpDate"))
                            myDR("p_transaction_quantity") = Items.Tables(1).Rows(i)("CLIDQty")
                            myDR("p_primary_quantity") = Items.Tables(1).Rows(i)("CLIDQty")
                            myDR("p_reason_code") = Items.Tables(1).Rows(i)("ReasonCode").ToString
                            myDR("p_transaction_reference") = Items.Tables(1).Rows(i)("DJ").ToString
                            myDR("o_return_status") = ""
                            myDR("o_return_message") = ""
                            ds.Tables(0).Rows.Add(myDR)
                        End If
                    Next
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                            ds.Tables(0).Rows(i).SetAdded()
                        End If
                    Next

                    Dim dsPost As DataSet = New DataSet
                    Try
                        dsPost = Post_SubInvTransfer(ds, LoginData, "PickOrder", TransactionID)
                    Catch ex As Exception
                        ErrorLogging("PickOrderPost-Post_SubInvTransfer", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                        Return ex.Message
                    End Try

                    DR = Nothing
                    DR = dsPost.Tables("transfer_table").Select("o_return_status = 'N' or o_return_status = ' ' or o_return_status IS Null ")
                    If DR.Length > 0 Then
                        PickOrderPost = DR(0)("o_return_message").ToString
                        If PickOrderPost = "" Then PickOrderPost = "No data is submitted in Oracle"
                        Return PickOrderPost
                    End If

                ElseIf SupplyType = "OSP" OrElse SupplyType = "FLOORSTOCK" Then     'No Data update in Oracle
                End If


                Try
                    Dim Sqlstr As String
                    Items.DataSetName = "Items"

                    Sqlstr = String.Format("exec sp_PickOrderPost '{0}','{1}',N'{2}'", LoginData.OrgCode, LoginData.User.ToUpper, DStoXML(Items))
                    PickOrderPost = da.ExecuteScalar(Sqlstr).ToString
                Catch ex As Exception
                    ErrorLogging("ProdPicking-sp_PickOrderPost", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                    PickOrderPost = "Data update error, please contact IT."
                End Try
                Return PickOrderPost

            Catch ex As Exception
                ErrorLogging("ProdPicking-PickOrderPost", LoginData.User.ToUpper, "PickOrder: " & PickOrder & "; " & ex.Message & ex.Source, "E")
                PickOrderPost = "Data update error, please contact IT."
            End Try
        End Using
    End Function

    Public Function OracleWIPIssue(ByVal ds As DataSet, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Dim DJ As String = ds.Tables("MatlList").Rows(0)("p_dj_name").ToString

            Try

                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                comm.ExecuteOracleNonQuery("")
                comm.Parameters.Clear()

                comm.CommandText = "apps.xxetr_wip_pkg.process_wip_issue"
                comm.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OrgID  'LoginData.OrgCode
                comm.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_item_num", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_item_rev", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_subinventory", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_locator", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_issue_quantity", OracleType.Double))
                comm.Parameters.Add(New OracleParameter("p_uom", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("p_reason", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("o_success_flag", OracleType.VarChar, 240))
                comm.Parameters.Add(New OracleParameter("o_error_msgg", OracleType.VarChar, 240))

                comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                comm.Parameters("o_error_msgg").Direction = ParameterDirection.InputOutput

                comm.Parameters("p_dj_name").SourceColumn = "p_dj_name"
                comm.Parameters("p_item_num").SourceColumn = "p_item_num"
                comm.Parameters("p_item_rev").SourceColumn = "p_item_rev"
                comm.Parameters("p_subinventory").SourceColumn = "p_subinventory"
                comm.Parameters("p_locator").SourceColumn = "p_locator"
                comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
                comm.Parameters("p_issue_quantity").SourceColumn = "p_issue_quantity"
                comm.Parameters("p_uom").SourceColumn = "p_uom"
                comm.Parameters("p_reason").SourceColumn = "p_reason"
                comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
                comm.Parameters("o_error_msgg").SourceColumn = "o_error_msgg"

                oda.InsertCommand = comm
                oda.Update(ds.Tables("MatlList"))

                Dim DR() As DataRow = Nothing
                DR = ds.Tables("MatlList").Select("o_success_flag = 'N' or o_success_flag = ' ' or o_success_flag IS Null ")
                If DR.Length = 0 Then   'No unsuccessful info is returned
                    comm.Transaction.Commit()
                    comm.Connection.Close()
                Else
                    comm.Transaction.Rollback()
                    comm.Connection.Close()

                    Dim i As Integer
                    For i = 0 To DR.Length - 1
                        Dim ErrMsg As String = ""
                        If DR(i)("o_error_msgg").ToString <> "" Then
                            'ErrMsg = "DJ " & DJ & " for item: " & DR(i)("p_item_num").ToString & " with IssuedQty " & DR(i)("p_issue_quantity").ToString & " and error message: " & DR(i)("o_error_msgg").ToString
                            ErrMsg = "DJ " & DJ & " for item: " & DR(i)("p_item_num").ToString & " with RTLot " & DR(i)("p_lot_number").ToString & ", Source SubInv/Locator " & DR(i)("p_subinventory").ToString & "/" & DR(i)("p_locator").ToString & ", IssuedQty " & DR(i)("p_issue_quantity").ToString & " and flag: " & DR(i)("o_success_flag").ToString & " and error message; " & DR(i)("o_error_msgg").ToString
                            ErrorLogging("ProdPicking-OracleWIPIssue", LoginData.User.ToUpper, ErrMsg, "I")
                        End If
                    Next
                End If

                Return ds  ' Return if there is error information returned from oracle

            Catch ex As Exception
                oda.InsertCommand.Transaction.Rollback()
                ErrorLogging("ProdPicking-OracleWIPIssue", LoginData.User.ToUpper, "DJ: " & DJ & ", " & ex.Message & ex.Source, "E")
                Return ds
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
        End Using

    End Function

    Public Function GetMOFromDN(ByVal DN As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            GetMOFromDN = New DataSet

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("MOrderList")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_dn_morders"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_dn_name", OracleType.VarChar, 50)).Value = DN

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "MOrderList")
                oda.SelectCommand.Connection.Close()

                GetMOFromDN = ds

                ds = New DataSet
                Dim Sqlstr As String

                Sqlstr = String.Format("Select * from T_Config with (nolock) where ConfigID = 'PIK002'")
                ds = da.ExecuteDataSet(Sqlstr, "Config")

                GetMOFromDN.Merge(ds.Tables("Config"))

            Catch ex As Exception
                ErrorLogging("Shipment-GetMOFromDN", LoginData.User, "DN: " & DN & "; " & ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function GetDNMOLines(ByVal DN As String, ByVal MO As String, ByVal PickStatus As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            'Dim PickStatus As String = "S"            'S: Released to Warehouse;   Y: Staged/Pick Confirmed
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ErrorMsg As String = "DN " & DN & " and MO " & MO

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim ds As DataSet = New DataSet()
                ds.Tables.Add("MatlList")
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.XXETR_wip_pkg.get_dn_morder_lines"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_move_order", OracleType.VarChar, 50)).Value = MO
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_dn_num", OracleType.VarChar, 50)).Value = DN
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_pick_status", OracleType.VarChar, 50)).Value = PickStatus
                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "MatlList")
                oda.SelectCommand.Connection.Close()

                Return ds
            Catch ex As Exception
                ErrorLogging("Shipment-GetDNMOLines", LoginData.User.ToUpper, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function

    Public Function ShipmentPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            ShipmentPost = ""
            Dim DN As String = Items.Tables(0).Rows(0)("delivery_number").ToString

            Try
                'Check if there has item over pick, prevent user doing over pick before post
                Dim PickDR() As DataRow = Nothing
                PickDR = Items.Tables(0).Select("PICKED_QTY > 0 and PICKED_QTY > REQUEST_QUANTITY ")
                If PickDR.Length > 0 Then
                    ShipmentPost = "Over pick is not allowed!"
                    Return ShipmentPost
                End If

                'check if exists CLID's status is 0, prevent user from slitting CLID after input, before post.
                Dim SqlCheck, StatusCode As String
                For Each dr As DataRow In Items.Tables(1).Rows
                    If dr("CLID").ToString = "" Then
                        Continue For
                    End If
                    SqlCheck = String.Format("select StatusCode from T_CLMaster where CLID='{0}'", dr("CLID").ToString)
                    StatusCode = Convert.ToString(da.ExecuteScalar(SqlCheck))
                    If StatusCode = "0" Then
                        Return "Exists invalid CLID: " & dr("CLID").ToString
                    End If
                Next

                Dim i As Integer
                Dim myDR As DataRow

                Try

                    Dim ds As DataSet = New DataSet()
                    ds.Tables.Add("MatlList")
                    ds.Tables(0).Columns.Add(New Data.DataColumn("TranHeaderID", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_dn_num", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_move_order", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_mo_line_id", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_revision", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_lot_number", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_quantity", System.Type.GetType("System.Decimal")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_subinv", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("p_locator", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_success_flag", System.Type.GetType("System.String")))
                    ds.Tables(0).Columns.Add(New Data.DataColumn("o_error_mssg", System.Type.GetType("System.String")))
                    ds.AcceptChanges()

                    For i = 0 To Items.Tables(0).Rows.Count - 1
                        If Items.Tables(0).Rows(i)("PICKED_QTY") > 0 Then
                            myDR = ds.Tables(0).NewRow
                            myDR("TranHeaderID") = Items.Tables(0).Rows(i)("TRANSACTION_HEADER_ID").ToString
                            myDR("p_dn_num") = Items.Tables(0).Rows(i)("DELIVERY_NUMBER").ToString
                            myDR("p_move_order") = Items.Tables(0).Rows(i)("MO_NUMBER").ToString
                            myDR("p_mo_line_id") = Items.Tables(0).Rows(i)("MO_LINE_ID").ToString
                            myDR("p_item_num") = Items.Tables(0).Rows(i)("ITEM_NUMBER").ToString
                            myDR("p_revision") = Items.Tables(0).Rows(i)("REVISION").ToString
                            myDR("p_lot_number") = Items.Tables(0).Rows(i)("LOT_NUMBER").ToString
                            myDR("p_subinv") = Items.Tables(0).Rows(i)("SOURCE_INV").ToString
                            myDR("p_locator") = Items.Tables(0).Rows(i)("SOURCE_LOC").ToString

                            Dim Qty As Decimal = Items.Tables(0).Rows(i)("PICKED_QTY")
                            myDR("p_quantity") = Math.Round(Qty, 6)                                   'Keep 6 decimals for Qty and send to Oracle
                            ds.Tables(0).Rows.Add(myDR)
                        End If
                    Next

                    ShipmentPost = ProcessMO(ds, LoginData)
                    If ShipmentPost <> "Y" Then Return ShipmentPost

                    'Insert data RM items / COO into Oracle customized table if there exists
                    If Items.Tables(2).Rows.Count > 0 Then
                        Dim RMItems As New DataSet
                        RMItems.Merge(Items.Tables(2))
                        InsertShipCOO(RMItems, LoginData)
                    End If

                Catch ex As Exception
                    ErrorLogging("ProdPicking-ShipmentPost1", LoginData.User, "DN: " & DN & "; " & ex.Message & ex.Source, "E")
                    Return "Data update error, please contact IT."
                End Try

                Try
                    Dim Sqlstr As String
                    Dim dsItem As DataSet = New DataSet

                    dsItem.DataSetName = "Items"
                    dsItem.Merge(Items.Tables(1))

                    Sqlstr = String.Format("exec sp_ShipmentPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, DN, LoginData.User.ToUpper, DStoXML(dsItem))
                    ShipmentPost = da.ExecuteScalar(Sqlstr).ToString
                Catch ex As Exception
                    ErrorLogging("ProdPicking-sp_ShipmentPost", LoginData.User, "DN: " & DN & "; " & ex.Message & ex.Source, "E")
                    ShipmentPost = "Data update error, please contact IT."
                End Try
                Return ShipmentPost

            Catch ex As Exception
                ErrorLogging("ProdPicking-ShipmentPost", LoginData.User, "DN: " & DN & "; " & ex.Message & ex.Source, "E")
                ShipmentPost = "Data update error, please contact IT."
            End Try
        End Using

    End Function

    Public Function InsertShipCOO(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Insert()
            Dim dsPick As New DataSet

            Dim DN, MO, PickStatus, Ermsg As String
            DN = p_ds.Tables(0).Rows(0)("DN").ToString
            MO = p_ds.Tables(0).Rows(0)("MO").ToString
            PickStatus = "Y"                             'S: Released to Warehouse;   Y: Staged/Pick Confirmed

            Try

                'Reading Staged/Pick Confirmed Move Order Lines for COO Qty assignment
                dsPick = GetDNMOLines(DN, MO, PickStatus, LoginData)
                If dsPick Is Nothing OrElse dsPick.Tables.Count = 0 Then
                    Ermsg = "Unknow error while reading data from Oracle for MO: " & MO
                    ErrorLogging("InsertShipCOO-GetDNMOLines", LoginData.User, Ermsg, "I")
                    Exit Function
                End If
                If dsPick.Tables(0).Columns.Count = 1 Then
                    Ermsg = dsPick.Tables(0).Rows(0)(0).ToString
                    ErrorLogging("InsertShipCOO-GetDNMOLines", LoginData.User, Ermsg, "I")
                    Exit Function
                End If

                dsPick.Merge(p_ds.Tables(0))

                p_ds = New DataSet
                p_ds = AssignCOOQty(dsPick, LoginData)
                If p_ds Is Nothing OrElse p_ds.Tables.Count = 0 OrElse p_ds.Tables(0).Rows.Count = 0 Then
                    Ermsg = "Unknow error while assign COO Qty for MO: " & MO
                    ErrorLogging("InsertShipCOO-AssignCOOQty", LoginData.User, Ermsg, "I")
                    Exit Function
                End If

                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_oe_pkg.ins_rm_shipment_coo"
                oda.InsertCommand.Parameters.Add("p_dn_number", OracleType.VarChar, 100)
                oda.InsertCommand.Parameters.Add("p_dn_detail_number", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_so_number", OracleType.VarChar, 100)
                oda.InsertCommand.Parameters.Add("p_so_line_number", OracleType.VarChar, 100)
                oda.InsertCommand.Parameters.Add("p_mo_number", OracleType.VarChar, 100)
                oda.InsertCommand.Parameters.Add("p_mo_line_number", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_inventory_item_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_coo", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_lot_number", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_quantity", OracleType.Double)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_message", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput

                oda.InsertCommand.Parameters("p_dn_number").SourceColumn = "DELIVERY_NUMBER"
                oda.InsertCommand.Parameters("p_dn_detail_number").SourceColumn = "DELIVERY_DETAIL_NUMBER"
                oda.InsertCommand.Parameters("p_so_number").SourceColumn = "SO_NUMBER"
                oda.InsertCommand.Parameters("p_so_line_number").SourceColumn = "SO_LINE_NUMBER"
                oda.InsertCommand.Parameters("p_mo_number").SourceColumn = "MO_NUMBER"
                oda.InsertCommand.Parameters("p_mo_line_number").SourceColumn = "MO_LINE_NUMBER"
                oda.InsertCommand.Parameters("p_inventory_item_id").SourceColumn = "INVENTORY_ITEM_ID"
                oda.InsertCommand.Parameters("p_coo").SourceColumn = "COO"
                oda.InsertCommand.Parameters("p_lot_number").SourceColumn = "LOT_NUMBER"
                oda.InsertCommand.Parameters("p_quantity").SourceColumn = "PICKED_QTY"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "Flag"
                oda.InsertCommand.Parameters("o_message").SourceColumn = "Message"

                oda.InsertCommand.Connection.Open()
                oda.Update(p_ds.Tables("COOData"))
                oda.InsertCommand.Connection.Close()

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables("COOData").Select(" Flag = 'N' or Flag = ' ' or Flag IS Null ")
                If DR.Length = 0 Then Return True

                'Record and Return Error message
                Dim i As Integer
                Dim ErrMsg As String = ""
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString <> "" Then
                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString
                    End If
                Next
                ErrorLogging("Shipment-InsertShipCOO", LoginData.User.ToUpper, ErrMsg, "I")
                Return False

            Catch ex As Exception
                ErrorLogging("Shipment-InsertShipCOO", LoginData.User.ToUpper, "MoveOrder: " & MO & "; " & ex.Message & ex.Source, "E")
                Return False
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

        End Using
    End Function

    Public Function AssignCOOQty(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As DataSet  ' For one MO has same items, assign COO Qty before save 

        Dim dsItem As New DataSet
        Dim COOData As DataTable
        COOData = Items.Tables(0).Clone
        COOData.TableName = "COOData"
        COOData.Columns.Add(New Data.DataColumn("COO", System.Type.GetType("System.String")))
        COOData.Columns.Add(New Data.DataColumn("Flag", System.Type.GetType("System.String")))
        COOData.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
        dsItem.Tables.Add(COOData)

        Try
            Dim i, j As Integer
            Dim myDR As DataRow

            For i = 0 To Items.Tables(1).Rows.Count - 1
                Dim Matl, LotNo As String
                Dim Qty As Decimal

                Matl = Items.Tables(1).Rows(i)("MaterialNo").ToString
                LotNo = Items.Tables(1).Rows(i)("RTLot").ToString
                Qty = Items.Tables(1).Rows(i)("PickedQty")

                Dim DR() As DataRow = Nothing
                'DR = Items.Tables(0).Select("ITEM_NUMBER='" & Matl & "' and LOT_NUMBER='" & LotNo & "'")
                DR = Items.Tables(0).Select("ITEM_NUMBER='" & Matl & "' and (LOT_NUMBER='" & LotNo & "' or LOT_NUMBER IS NULL) ")
                If DR.Length > 0 Then
                    For j = 0 To DR.Length - 1
                        If DR(j)("REQUEST_QUANTITY") = 0 Then Continue For

                        If DR(j)("REQUEST_QUANTITY") >= Qty Then
                            DR(j)("PICKED_QTY") = Qty
                            DR(j)("REQUEST_QUANTITY") = DR(j)("REQUEST_QUANTITY") - Qty

                            myDR = dsItem.Tables("COOData").NewRow              'record the CLID
                            myDR.ItemArray = DR(j).ItemArray
                            myDR("COO") = Items.Tables(1).Rows(i)("COO").ToString
                            dsItem.Tables("COOData").Rows.Add(myDR)
                            Exit For
                        ElseIf DR(j)("REQUEST_QUANTITY") < Qty And j < DR.Length - 1 Then
                            DR(j)("PICKED_QTY") = DR(j)("REQUEST_QUANTITY")
                            Qty = Qty - DR(j)("REQUEST_QUANTITY")
                            DR(j)("REQUEST_QUANTITY") = 0
                        Else                                            'j=DR.Length-1
                            DR(j)("PICKED_QTY") = Qty
                            DR(j)("REQUEST_QUANTITY") = DR(j)("REQUEST_QUANTITY") - Qty
                            If DR(j)("REQUEST_QUANTITY") < 0 Then DR(j)("REQUEST_QUANTITY") = 0
                        End If

                        myDR = dsItem.Tables("COOData").NewRow              'record the CLID
                        myDR.ItemArray = DR(j).ItemArray
                        myDR("COO") = Items.Tables(1).Rows(i)("COO").ToString
                        dsItem.Tables("COOData").Rows.Add(myDR)
                    Next
                    Items.Tables(0).AcceptChanges()
                End If
            Next

            Return dsItem

        Catch ex As Exception
            ErrorLogging("Shipment-AssignCOOQty", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    'Public Function ProcessMO(ByVal ds As DataSet, ByVal LoginData As ERPLogin) As String
    '    Using da As DataAccess = GetDataAccess()

    '        Try
    '            Dim DN As String = ds.Tables(0).Rows(0)("p_dn_num").ToString
    '            Dim MO As String = ds.Tables(0).Rows(0)("p_move_order").ToString

    '            Dim ErrorMsg As String = "MO "
    '            Dim ModuleName As String = "ProdPicking-ProcessMO"
    '            If DN <> "" Then
    '                ErrorMsg = "DN " & DN & " and MO "
    '                ModuleName = "ProdPicking-ProcessShipment"
    '            End If

    '            Dim oda As New OracleDataAdapter
    '            Dim comm As New OracleCommand

    '            Dim OrgID As String = GetOrgID(LoginData.OrgCode)
    '            'oda = da.Oda_Update_Tran
    '            'comm = da.Ora_Command_Trans()

    '            Try
    '                oda = da.Oda_Update_Tran
    '                comm = da.Ora_Command_Trans()
    '            Catch ex As Exception
    '                ErrorLogging("ProcessMO-OpenException", LoginData.User, ErrorMsg & MO & ", " & ex.Message & ex.Source, "E")
    '                ProcessMO = ex.Message
    '                Return ProcessMO
    '                Exit Function
    '            End Try

    '            Try
    '                comm.CommandType = CommandType.StoredProcedure
    '                comm.CommandText = "apps.XXETR_wip_pkg.initialize"
    '                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
    '                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
    '                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
    '                comm.ExecuteOracleNonQuery("")
    '                comm.Parameters.Clear()

    '                'If DN <> "" Then
    '                '    comm.CommandText = "apps.XXETR_wip_pkg.transact_move_order"
    '                '    comm.Parameters.Add(New OracleParameter("p_mo_line_id", OracleType.Int32))
    '                '    comm.Parameters.Add(New OracleParameter("p_revision", OracleType.VarChar, 50))
    '                '    comm.Parameters("p_mo_line_id").SourceColumn = "p_mo_line_id"
    '                '    comm.Parameters("p_revision").SourceColumn = "p_revision"
    '                'Else
    '                '    comm.CommandText = "apps.XXETR_wip_pkg.process_move_order"
    '                'End If
    '                comm.CommandText = "apps.XXETR_wip_pkg.transact_move_order"
    '                comm.Parameters.Add(New OracleParameter("p_mo_line_id", OracleType.Int32))
    '                comm.Parameters.Add(New OracleParameter("p_revision", OracleType.VarChar, 50))
    '                comm.Parameters("p_mo_line_id").SourceColumn = "p_mo_line_id"
    '                comm.Parameters("p_revision").SourceColumn = "p_revision"

    '                comm.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32)).Value = CInt(LoginData.UserID)
    '                comm.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 20)).Value = OrgID  'LoginData.OrgCode
    '                comm.Parameters.Add(New OracleParameter("p_move_order", OracleType.VarChar, 50))
    '                comm.Parameters.Add(New OracleParameter("p_item_num", OracleType.VarChar, 20))
    '                comm.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 50))
    '                comm.Parameters.Add(New OracleParameter("p_quantity", OracleType.Double))
    '                comm.Parameters.Add(New OracleParameter("p_subinv", OracleType.VarChar, 50))
    '                comm.Parameters.Add(New OracleParameter("p_locator", OracleType.VarChar, 50))
    '                comm.Parameters.Add(New OracleParameter("o_success_flag", OracleType.VarChar, 10))
    '                comm.Parameters.Add(New OracleParameter("o_error_mssg", OracleType.VarChar, 2000))

    '                comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
    '                comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

    '                comm.Parameters("p_move_order").SourceColumn = "p_move_order"
    '                comm.Parameters("p_item_num").SourceColumn = "p_item_num"
    '                comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
    '                comm.Parameters("p_quantity").SourceColumn = "p_quantity"
    '                comm.Parameters("p_subinv").SourceColumn = "p_subinv"
    '                comm.Parameters("p_locator").SourceColumn = "p_locator"
    '                comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
    '                comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

    '                oda.InsertCommand = comm
    '                oda.Update(ds.Tables("MatlList"))

    '                If ds.Tables("MatlList").Rows.Count = 0 Then
    '                    oda.InsertCommand.Transaction.Rollback()
    '                    oda.InsertCommand.Connection.Close()
    '                    ProcessMO = "No data is submitted in Oracle"
    '                    Return ProcessMO
    '                End If

    '                Dim DR() As DataRow = Nothing
    '                DR = ds.Tables("MatlList").Select(" o_success_flag = 'N' or o_success_flag = ' ' or o_success_flag IS Null ")
    '                If DR.Length = 0 Then
    '                    oda.InsertCommand.Transaction.Commit()
    '                    oda.InsertCommand.Connection.Close()
    '                    ProcessMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
    '                    Return ProcessMO
    '                End If


    '                'Rollback and record error message if Oracle Data update failed
    '                oda.InsertCommand.Transaction.Rollback()
    '                oda.InsertCommand.Connection.Close()

    '                Dim i As Integer
    '                For i = 0 To DR.Length - 1
    '                    Dim ErrMsg As String = ""
    '                    MO = DR(i)("p_move_order").ToString
    '                    ErrMsg = ErrorMsg & MO & " and TranHeaderID " & DR(i)("TranHeaderID").ToString & " for item: " & DR(i)("p_item_num").ToString & " with RTLot " & DR(i)("p_lot_number").ToString & ", Source SubInv/Locator " & DR(i)("p_subinv").ToString & "/" & DR(i)("p_locator").ToString & ", IssuedQty " & DR(i)("p_quantity").ToString & " and flag: " & DR(i)("o_success_flag").ToString & " and error message; " & DR(i)("o_error_mssg").ToString
    '                    ErrorLogging(ModuleName, LoginData.User, ErrMsg, "I")
    '                Next

    '                MO = DR(0)("p_move_order").ToString
    '                ProcessMO = DR(0)("o_error_mssg").ToString
    '                If ProcessMO = "" Then
    '                    ProcessMO = "No data is submitted in Oracle"
    '                    Return ProcessMO
    '                End If

    '                If ProcessMO.ToUpper.Contains("ORA-00054") Then
    '                    ProcessMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with item: " & DR(0)("p_item_num").ToString
    '                Else
    '                    ProcessMO = ErrorMsg & MO & " for item: " & DR(0)("p_item_num").ToString & " with error message; " & DR(0)("o_error_mssg").ToString
    '                End If
    '                Return ProcessMO                                'Return Error message

    '            Catch ex As Exception
    '                ErrorLogging(ModuleName, LoginData.User, ErrorMsg & MO & ", " & ex.Message & ex.Source, "E")
    '                oda.InsertCommand.Transaction.Rollback()
    '                ProcessMO = ex.Message
    '            Finally
    '                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
    '                If oda.UpdateCommand.Connection.State <> ConnectionState.Closed Then oda.UpdateCommand.Connection.Close()
    '                If oda.DeleteCommand.Connection.State <> ConnectionState.Closed Then oda.DeleteCommand.Connection.Close()
    '                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
    '            End Try

    '        Catch ex As Exception
    '            ErrorLogging("ProcessMO-Exception", LoginData.User, ex.Message & ex.Source, "E")
    '            ProcessMO = ex.Message
    '        End Try

    '    End Using

    'End Function
    Public Function ProcessMO(ByVal ds As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim DN As String = ds.Tables(0).Rows(0)("p_dn_num").ToString
                Dim MO As String = ds.Tables(0).Rows(0)("p_move_order").ToString

                Dim ErrorMsg As String = "MO "
                Dim ModuleName As String = "ProdPicking-ProcessMO"
                If DN <> "" Then
                    ErrorMsg = "DN " & DN & " and MO "
                    ModuleName = "ProdPicking-ProcessShipment"
                End If

                Dim oda As OracleDataAdapter
                Dim comm As OracleCommand

                Dim OrgID As String = GetOrgID(LoginData.OrgCode)
                'oda = da.Oda_Update_Tran
                'comm = da.Ora_Command_Trans()

                Try
                    oda = New OracleDataAdapter
                    comm = da.Ora_Command_Trans()
                Catch ex As Exception
                    ErrorLogging("ProcessMO-OpenException", LoginData.User, ErrorMsg & MO & ", " & ex.Message & ex.Source, "E")
                    ProcessMO = ex.Message
                    Return ProcessMO
                    Exit Function
                End Try

                Try
                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.XXETR_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                    comm.ExecuteOracleNonQuery("")
                    comm.Parameters.Clear()

                    '    comm.CommandText = "apps.XXETR_wip_pkg.process_move_order"
                    comm.CommandText = "apps.XXETR_wip_pkg.transact_move_order"
                    comm.Parameters.Add("p_mo_line_id", OracleType.Int32)
                    comm.Parameters.Add("p_revision", OracleType.VarChar, 50)
                    comm.Parameters("p_mo_line_id").SourceColumn = "p_mo_line_id"
                    comm.Parameters("p_revision").SourceColumn = "p_revision"

                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                    comm.Parameters.Add("p_org_code", OracleType.VarChar, 20).Value = OrgID  'LoginData.OrgCode
                    comm.Parameters.Add("p_move_order", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_lot_number", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_quantity", OracleType.Double)
                    comm.Parameters.Add("p_subinv", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_locator", OracleType.VarChar, 50)
                    comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
                    comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

                    comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                    comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                    comm.Parameters("p_move_order").SourceColumn = "p_move_order"
                    comm.Parameters("p_item_num").SourceColumn = "p_item_num"
                    comm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
                    comm.Parameters("p_quantity").SourceColumn = "p_quantity"
                    comm.Parameters("p_subinv").SourceColumn = "p_subinv"
                    comm.Parameters("p_locator").SourceColumn = "p_locator"
                    comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
                    comm.Parameters("o_error_mssg").SourceColumn = "o_error_mssg"

                    oda.InsertCommand = comm
                    oda.Update(ds.Tables("MatlList"))

                    If ds.Tables("MatlList").Rows.Count = 0 Then
                        comm.Transaction.Rollback()
                        comm.Connection.Close()
                        ProcessMO = "No data is submitted in Oracle"
                        Return ProcessMO
                    End If

                    Dim DR() As DataRow = Nothing
                    DR = ds.Tables("MatlList").Select(" o_success_flag = 'N' or o_success_flag = ' ' or o_success_flag IS Null ")
                    If DR.Length = 0 Then
                        comm.Transaction.Commit()
                        comm.Connection.Close()
                        ProcessMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
                        Return ProcessMO
                    End If

                    'Rollback and record error message if Oracle Data update failed
                    comm.Transaction.Rollback()
                    comm.Connection.Close()

                    Dim i As Integer
                    For i = 0 To DR.Length - 1
                        Dim ErrMsg As String = ""
                        MO = DR(i)("p_move_order").ToString
                        ErrMsg = ErrorMsg & MO & " and TranHeaderID " & DR(i)("TranHeaderID").ToString & " for item: " & DR(i)("p_item_num").ToString & " with RTLot " & DR(i)("p_lot_number").ToString & ", Source SubInv/Locator " & DR(i)("p_subinv").ToString & "/" & DR(i)("p_locator").ToString & ", IssuedQty " & DR(i)("p_quantity").ToString & " and flag: " & DR(i)("o_success_flag").ToString & " and error message; " & DR(i)("o_error_mssg").ToString
                        ErrorLogging(ModuleName, LoginData.User, ErrMsg, "I")
                    Next

                    MO = DR(0)("p_move_order").ToString
                    ProcessMO = DR(0)("o_error_mssg").ToString
                    If ProcessMO = "" Then
                        ProcessMO = "No data is submitted in Oracle"
                        Return ProcessMO
                    End If

                    If ProcessMO.ToUpper.Contains("ORA-00054") Then
                        ProcessMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with item: " & DR(0)("p_item_num").ToString
                    Else
                        ProcessMO = ErrorMsg & MO & " for item: " & DR(0)("p_item_num").ToString & " with error message; " & DR(0)("o_error_mssg").ToString
                    End If
                    Return ProcessMO                                'Return Error message

                Catch ex As Exception
                    ErrorLogging(ModuleName, LoginData.User, ErrorMsg & MO & ", " & ex.Message & ex.Source, "E")
                    comm.Transaction.Rollback()
                    ProcessMO = ex.Message
                Finally
                    If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
                End Try

            Catch ex As Exception
                ErrorLogging("ProcessMO-Exception", LoginData.User, ex.Message & ex.Source, "E")
                ProcessMO = ex.Message
            End Try

        End Using

    End Function

    Public Function TransactMO(ByVal ds As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim s As String = ""
            Dim OpenTime, UpdateTime As String

            Try
                s = 1
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)
                Dim MO As String = ds.Tables(0).Rows(0)("p_move_order").ToString

                s = 2
                Dim sda As SqlClient.SqlDataAdapter = da.Sda_Insert()

                Try

                    'Dim Sqlstr As String
                    'Sqlstr = String.Format("exec ora_initialize  '{0}','{1}','{2}' ", CInt(LoginData.UserID), CInt(LoginData.RespID_Inv), CInt(LoginData.AppID_Inv))
                    'da.ExecuteScalar(Sqlstr)

                    s = 3 & ", " & GetServerDate.ToString
                    sda.InsertCommand.CommandType = CommandType.StoredProcedure
                    sda.InsertCommand.CommandText = "ora_Transact_move_order"
                    sda.InsertCommand.CommandTimeout = TimeOut_M5

                    sda.InsertCommand.Parameters.Add("@p_user_id", SqlDbType.Int).Value = CInt(LoginData.UserID)
                    'sda.InsertCommand.Parameters.Add("@p_resp_id", SqlDbType.Int).Value = CInt(LoginData.RespID_Inv)   '54050
                    'sda.InsertCommand.Parameters.Add("@p_appl_id", SqlDbType.Int).Value = CInt(LoginData.AppID_Inv)
                    sda.InsertCommand.Parameters.Add("@p_org_code", SqlDbType.VarChar, 20).Value = OrgID  'LoginData.OrgCode
                    sda.InsertCommand.Parameters.Add("@p_move_order", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@p_mo_line_id", SqlDbType.Int)
                    sda.InsertCommand.Parameters.Add("@p_item_num", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@p_lot_number", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@p_quantity", SqlDbType.Float)
                    sda.InsertCommand.Parameters.Add("@p_subinv", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@p_locator", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@p_revision", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@o_success_flag", SqlDbType.VarChar, 2)
                    sda.InsertCommand.Parameters.Add("@o_error_mssg", SqlDbType.VarChar, 4000)

                    sda.InsertCommand.Parameters("@p_move_order").SourceColumn = "p_move_order"
                    sda.InsertCommand.Parameters("@p_mo_line_id").SourceColumn = "p_mo_line_id"
                    sda.InsertCommand.Parameters("@p_item_num").SourceColumn = "p_item_num"
                    sda.InsertCommand.Parameters("@p_lot_number").SourceColumn = "p_lot_number"
                    sda.InsertCommand.Parameters("@p_quantity").SourceColumn = "p_quantity"
                    sda.InsertCommand.Parameters("@p_subinv").SourceColumn = "p_subinv"
                    sda.InsertCommand.Parameters("@p_locator").SourceColumn = "p_locator"
                    sda.InsertCommand.Parameters("@p_revision").SourceColumn = "p_revision"
                    sda.InsertCommand.Parameters("@o_success_flag").SourceColumn = "o_success_flag"
                    sda.InsertCommand.Parameters("@o_error_mssg").SourceColumn = "o_error_mssg"

                    sda.InsertCommand.Parameters("@o_success_flag").Direction = ParameterDirection.Output
                    sda.InsertCommand.Parameters("@o_error_mssg").Direction = ParameterDirection.Output

                    OpenTime = GetServerDate.ToString
                    s = 4 & ", OpenedOn: " & OpenTime

                    Try
                        sda.InsertCommand.Connection.Open()
                    Catch ex As Exception
                        ErrorLogging("TransactMO-OpenException", LoginData.User & ", s=" & s, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                        TransactMO = ex.Message
                        Return TransactMO
                        Exit Function
                    End Try

                    UpdateTime = GetServerDate.ToString
                    s = 5 & ", OpenedOn: " & OpenTime & ", UpdatedOn: " & UpdateTime
                    sda.Update(ds.Tables("MatlList"))
                    sda.InsertCommand.Connection.Close()

                Catch ex As Exception
                    ErrorLogging("TransactMO-ora_Transact_move_order", LoginData.User & ", s=" & s, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                    TransactMO = "Transact MO error"
                Finally
                    If sda.InsertCommand.Connection.State <> ConnectionState.Closed Then sda.InsertCommand.Connection.Close()
                End Try

                s = 6
                If ds.Tables("MatlList").Rows.Count = 0 Then
                    TransactMO = "No data is submitted in Oracle"
                    Return TransactMO
                End If

                s = 7
                Dim DR() As DataRow = Nothing
                DR = ds.Tables("MatlList").Select(" o_success_flag = 'N' or o_success_flag = ' ' or o_success_flag IS Null ")
                If DR.Length = 0 Then
                    TransactMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
                    Return TransactMO
                End If

                s = 8 & ", OpenedOn: " & OpenTime & ", UpdatedOn: " & UpdateTime

                Dim i As Integer
                For i = 0 To DR.Length - 1
                    Dim ErrMsg As String = ""
                    MO = DR(i)("p_move_order").ToString
                    ErrMsg = "MO " & MO & " and TranHeaderID " & DR(i)("TranHeaderID").ToString & " for item: " & DR(i)("p_item_num").ToString & " with RTLot " & DR(i)("p_lot_number").ToString & ", Source SubInv/Locator " & DR(i)("p_subinv").ToString & "/" & DR(i)("p_locator").ToString & ", IssuedQty " & DR(i)("p_quantity").ToString & " and flag: " & DR(i)("o_success_flag").ToString & " and error message; " & DR(i)("o_error_mssg").ToString
                    ErrorLogging("ProdPicking-TransactMO", LoginData.User & ", s=" & s, ErrMsg, "I")
                Next

                s = 9
                MO = DR(0)("p_move_order").ToString
                TransactMO = DR(0)("o_error_mssg").ToString
                If TransactMO = "" Then
                    TransactMO = "No data is submitted in Oracle"
                    Return TransactMO
                End If

                s = 10
                If TransactMO.ToUpper.Contains("ORA-00054") Then
                    TransactMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with item: " & DR(0)("p_item_num").ToString
                Else
                    TransactMO = "MO " & MO & " for item: " & DR(0)("p_item_num").ToString & " with error message; " & DR(0)("o_error_mssg").ToString
                End If
                Return TransactMO                                'Return Error message

            Catch ex As Exception
                ErrorLogging("TransactMO-Exception", LoginData.User & ", s=" & s, ex.Message & ex.Source, "E")
                TransactMO = ex.Message
            End Try

        End Using

    End Function

    Public Function GetTranHeaderID(ByVal LoginData As ERPLogin, ByVal TypeID As String) As String
        Dim NextID As String = ""
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra, k As Integer

        Try

            myConn.Open()
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            'myCommand.Parameters(0).Size = 9
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", TypeID)
            If TypeID = "HDID" Then
                myCommand.Parameters(0).Size = 9
            ElseIf TypeID = "PLID" Then
                myCommand.Parameters(0).Size = 10
            End If

            'Try up to 5 times when failed getting next id
            NextID = ""
            k = 0
            While (k < 5 And NextID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("ProdPicking-GetTranHeaderID", "Deadlocked? " & Str(k), "Failed to get next TranHeaderID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextID
        Catch ex As Exception
            ErrorLogging("ProdPicking-GetTranHeaderID", LoginData.User, ex.Message & ex.Source, "E")
            Return NextID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetDNPickedLists(ByVal LoginData As ERPLogin, ByVal DN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim dsItem As DataSet = New DataSet

            Dim ErrorMsg As String = "DN " & DN & "; "

            Try

                Dim Sqlstr As String
                Sqlstr = String.Format("Select MaterialNo, Sum(QtyBaseUOM) As PickedQty, ReverseQty = 0.000 from T_CLMaster with (nolock) GROUP BY OrgCode, ShipmentNo, StatusCode, MaterialNo having StatusCode = 0 and OrgCode = '{0}' and ShipmentNo = '{1}' ", LoginData.OrgCode, DN)
                dsItem = da.ExecuteDataSet(Sqlstr, "MatlList")
                If dsItem Is Nothing OrElse dsItem.Tables.Count = 0 OrElse dsItem.Tables(0).Rows.Count = 0 Then
                    Return dsItem
                End If

                Dim ds As DataSet = New DataSet()
                Sqlstr = String.Format("Select * from T_Config with (nolock) where ConfigID = 'PIK002'")
                ds = da.ExecuteDataSet(Sqlstr, "Config")

                dsItem.Merge(ds.Tables("Config"))
                Return dsItem

            Catch ex As Exception
                ErrorLogging("ProdPicking-GetDNPickedLists", LoginData.User.ToUpper, ErrorMsg & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using
    End Function

    Public Function GetDNPickedCLIDs(ByVal LoginData As ERPLogin, ByVal DN As String, ByVal CLID As String) As DataSet

        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs As DataSet = New DataSet

            Try
                Dim Sqlstr, ErrMsg As String
                If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then 'Box ID
                    Sqlstr = String.Format("Select TOP (1) OrgCode,CLID,QtyBaseUOM,BoxID,MaterialNo,MaterialRevision,BaseUOM,SLOC,StorageBin,CONVERT(varchar, ExpDate, 101) as Expdate,RTLot, ShipmentNo, DestSubInv='', DestLocator='', StorageType='', Status = '', Message = '' from T_CLMaster with (nolock) where OrgCode = '{0}' and BoxID = '{1}'", LoginData.OrgCode, CLID)
                Else
                    Sqlstr = String.Format("Select OrgCode,CLID,QtyBaseUOM,BoxID,MaterialNo,MaterialRevision,BaseUOM,SLOC,StorageBin,CONVERT(varchar, ExpDate, 101) as Expdate,RTLot, ShipmentNo, DestSubInv='', DestLocator='', StorageType='', Status = '', Message = '' from T_CLMaster with (nolock) where OrgCode = '{0}' and CLID = '{1}'", LoginData.OrgCode, CLID)
                End If
                myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDS")
                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 OrElse myCLIDs.Tables(0).Rows.Count = 0 Then
                    Return myCLIDs
                End If

                If myCLIDs.Tables(0).Rows(0)("ShipmentNo") Is DBNull.Value Or myCLIDs.Tables(0).Rows(0)("ShipmentNo").ToString <> DN Then
                    ErrMsg = "CLID " & CLID & " didn't issued to DN " & DN & ", please check."
                    myCLIDs.Tables(0).Rows(0)("Message") = ErrMsg
                    Return myCLIDs
                End If

                'Slot is available, only read the scanned CLID 
                If LoginData.ResetFlag = True Then
                    Return myCLIDs
                End If

                Dim MaterialNo As String
                MaterialNo = myCLIDs.Tables(0).Rows(0)("MaterialNo").ToString
                Sqlstr = String.Format("Select OrgCode,CLID,QtyBaseUOM,BoxID,MaterialNo,MaterialRevision,BaseUOM,SLOC,StorageBin,CONVERT(varchar, ExpDate, 101) as Expdate,RTLot, ShipmentNo, DestSubInv='', DestLocator='', StorageType='', Status = '', Message = '' from T_CLMaster with (nolock) where StatusCode = 0 and OrgCode = '{0}' and ShipmentNo = '{1}' and MaterialNo = '{2}'", LoginData.OrgCode, DN, MaterialNo)

                myCLIDs = New DataSet
                myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDS")
                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("ProdPicking-GetDNPickedCLIDs", LoginData.User.ToUpper, "DN: " & DN & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function SOReversalPost(ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim DN As String = Items.Tables(0).Rows(0)("ShipmentNo").ToString
            Dim MaterialNo As String = Items.Tables(0).Rows(0)("MaterialNo").ToString
            Dim ErrorMsg As String = "Shipment Reversal for DN " & DN & " and item " & MaterialNo & " with error: "

            Try

                Dim DR() As DataRow
                Dim ds As DataSet = New DataSet
                Dim CLIDS As DataTable = Items.Tables(0).Clone()
                ds.Tables.Add(CLIDS)

                Dim i As Integer
                Dim myDR As DataRow

                Try
                    For i = 0 To Items.Tables(0).Rows.Count - 1
                        Dim Rev, SubInv, Locator, RTLot As String
                        Rev = Items.Tables(0).Rows(i)("MaterialRevision").ToString
                        SubInv = Items.Tables(0).Rows(i)("SLOC").ToString
                        Locator = Items.Tables(0).Rows(i)("StorageBin").ToString
                        RTLot = Items.Tables(0).Rows(i)("RTLot").ToString

                        Dim SqlStr As String
                        If Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                            SqlStr = "MaterialRevision='" & Rev & "' and SLOC = '" & SubInv & "' and StorageBin = '" & Locator & "' and RTLot = '" & RTLot & "'"
                        Else
                            SqlStr = "MaterialRevision='" & Rev & "' and SLOC = '" & SubInv & "' and StorageBin = '" & Locator & "' and RTLot = '" & RTLot & "' and ExpDate = '" & Items.Tables(0).Rows(i)("ExpDate") & "'"
                        End If

                        DR = Nothing
                        DR = ds.Tables(0).Select(SqlStr)
                        If DR.Length = 0 Then
                            myDR = ds.Tables(0).NewRow()
                            myDR("MaterialNo") = Items.Tables(0).Rows(i)("MaterialNo")
                            myDR("MaterialRevision") = Items.Tables(0).Rows(i)("MaterialRevision").ToString
                            myDR("QtyBaseUOM") = Items.Tables(0).Rows(i)("QtyBaseUOM")
                            myDR("BaseUOM") = Items.Tables(0).Rows(i)("BaseUOM")
                            myDR("SLOC") = Items.Tables(0).Rows(i)("SLOC").ToString
                            myDR("StorageBin") = Items.Tables(0).Rows(i)("StorageBin").ToString
                            myDR("RTLot") = Items.Tables(0).Rows(i)("RTLot").ToString
                            myDR("ExpDate") = Items.Tables(0).Rows(i)("ExpDate")
                            myDR("DestSubInv") = Items.Tables(0).Rows(i)("DestSubInv")
                            myDR("DestLocator") = Items.Tables(0).Rows(i)("DestLocator")
                            myDR("ShipmentNo") = Items.Tables(0).Rows(i)("ShipmentNo")
                            myDR("Status") = ""
                            myDR("Message") = ""
                            ds.Tables(0).Rows.Add(myDR)
                        Else
                            Dim Qty As Decimal
                            Qty = DR(0)("QtyBaseUOM") + Items.Tables(0).Rows(i)("QtyBaseUOM")

                            DR(0)("QtyBaseUOM") = Math.Round(Qty, 5)
                            DR(0).AcceptChanges()
                            DR(0).SetAdded()
                        End If
                    Next
                Catch ex As Exception
                    ErrorLogging("ProdPicking-SOReversalPost1", LoginData.User, ErrorMsg & ex.Message & ex.Source, "E")
                    Return "Shipment Reversal Transfer error"
                End Try

                Try
                    SOReversalPost = SOReversal_Transfer(LoginData, ds)
                    If SOReversalPost <> "Y" Then Exit Function
                Catch ex As Exception
                    ErrorLogging("ProdPicking-SOReversalPost2", LoginData.User.ToUpper, ErrorMsg & ex.Message & ex.Source, "E")
                    Return "Shipment Reversal Transfer error"
                End Try


                Try
                    Dim Sqlstr As String
                    Items.DataSetName = "Items"

                    Sqlstr = String.Format("exec sp_SOReversalPost '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, DN, LoginData.User.ToUpper, DStoXML(Items))
                    SOReversalPost = da.ExecuteScalar(Sqlstr).ToString
                Catch ex As Exception
                    ErrorLogging("ProdPicking-sp_SOReversalPost", LoginData.User, "DN: " & DN & "; " & ex.Message & ex.Source, "E")
                    SOReversalPost = "Data update error, please contact IT."
                End Try
                Return SOReversalPost

            Catch ex As Exception
                ErrorLogging("ProdPicking-SOReversalPost", LoginData.User, ErrorMsg & ex.Message & ex.Source, "E")
                SOReversalPost = "Data update error, please contact IT."
            End Try

        End Using

    End Function

    Public Function SOReversal_Transfer(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Dim DN As String = p_ds.Tables(0).Rows(0)("ShipmentNo").ToString
            Dim MaterialNo As String = p_ds.Tables(0).Rows(0)("MaterialNo").ToString

            Dim ErrorMsg As String = "Shipment Reversal Transfer for DN " & DN & " and item " & MaterialNo

            Dim aa As OracleString
            Dim oda As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Try

                Dim TransactionID As Long
                TransactionID = CLng(GetTranHeaderID(LoginData, "HDID"))

                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                comm.ExecuteOracleNonQuery(aa)
                comm.Parameters.Clear()

                comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.subinventory_transfer"
                comm.Parameters.Add("p_timeout", OracleType.Int32).Value = 1000 * 60 * 30
                comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID   'LoginData.OrgCode
                comm.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                'comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240).Value = ""
                'comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240).Value = ""
                comm.Parameters.Add("p_user_id", OracleType.Double).Value = LoginData.UserID
                comm.Parameters.Add("p_source_line_id", OracleType.Double).Value = LoginData.UserID
                comm.Parameters.Add("p_source_header_id", OracleType.Double).Value = LoginData.UserID
                comm.Parameters.Add("p_transaction_uom", OracleType.VarChar, 240)
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

                comm.Parameters.Add("o_return_status", OracleType.VarChar, 240)
                comm.Parameters.Add("o_return_message", OracleType.VarChar, 800)
                comm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
                comm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

                comm.Parameters("p_transaction_uom").SourceColumn = "BaseUOM"
                comm.Parameters("p_item_segment1").SourceColumn = "MaterialNo"
                comm.Parameters("p_item_revision").SourceColumn = "MaterialRevision"
                comm.Parameters("p_subinventory_source").SourceColumn = "SLOC"
                comm.Parameters("p_locator_source").SourceColumn = "StorageBin"
                comm.Parameters("p_subinventory_destination").SourceColumn = "DestSubInv"
                comm.Parameters("p_locator_destination").SourceColumn = "DestLocator"
                comm.Parameters("p_lot_number").SourceColumn = "RTLot"
                comm.Parameters("p_lot_expiration_date").SourceColumn = "Expdate"
                comm.Parameters("p_transaction_quantity").SourceColumn = "QtyBaseUOM"
                comm.Parameters("p_primary_quantity").SourceColumn = "QtyBaseUOM"
                comm.Parameters("o_return_status").SourceColumn = "Status"
                comm.Parameters("o_return_message").SourceColumn = "Message"

                oda.InsertCommand = comm
                oda.Update(p_ds.Tables(0))

                Dim Flag, Msg As String
                Flag = comm.Parameters("o_return_status").Value.ToString
                Msg = comm.Parameters("o_return_message").Value.ToString

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables(0).Select("Status = 'N' or Status = ' '")
                If DR.Length = 0 Then
                    comm.Transaction.Commit()
                    comm.Connection.Close()
                    Return DirectCast(Flag, String)
                Else
                    Dim ErrMsg As String = ErrorMsg & " with error message: " & Msg
                    ErrorLogging("ProdPicking-SOReversal_Transfer1", LoginData.User.ToUpper, ErrMsg, "I")
                    comm.Transaction.Rollback()
                    comm.Connection.Close()

                    If Msg.Contains("Available quantity is not enough") Then
                        Msg = Msg & " Make sure the item has done backorder in Oracle first."
                    End If
                    Return DirectCast(Msg, String)
                End If

            Catch ex As Exception
                ErrorLogging("ProdPicking-SOReversal_Transfer", LoginData.User.ToUpper, ErrorMsg & ", " & ex.Message & ex.Source, "E")
                comm.Transaction.Rollback()
                Return "Shipment Reversal Transfer error"
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
        End Using

    End Function



    Public Function DJReadProcessLabel(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Insert()

            Dim ErrorMsg As String = "Expand BOM with error for ProcessLabel; "

            Try
                Dim dtModel As New DataTable
                dtModel = myMOData.Items.Tables(0).Copy

                Dim dtMatl As New DataTable
                dtMatl = myMOData.Items.Tables(1).Copy
                myMOData.Items = New DataSet

                Dim dtDJ As DataTable
                dtDJ = New Data.DataTable("DJ")
                dtDJ.Columns.Add(New Data.DataColumn("p_dj_name", System.Type.GetType("System.String")))
                dtDJ.Columns.Add(New Data.DataColumn("o_identity", System.Type.GetType("System.String")))
                dtDJ.Columns.Add(New Data.DataColumn("o_flag", System.Type.GetType("System.String")))
                dtDJ.Columns.Add(New Data.DataColumn("o_err_msg", System.Type.GetType("System.String")))

                Dim dtItem As DataTable
                dtItem = New Data.DataTable("Item")
                dtItem.Columns.Add(New Data.DataColumn("p_identity", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("p_dj_name", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("p_com_item", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_floor", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_dj_start_qty", System.Type.GetType("System.Decimal")))
                dtItem.Columns.Add(New Data.DataColumn("o_dj_model", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_assy", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_req_qty", System.Type.GetType("System.Decimal")))
                dtItem.Columns.Add(New Data.DataColumn("o_usage_qty", System.Type.GetType("System.Decimal")))
                dtItem.Columns.Add(New Data.DataColumn("o_forming_code", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_refdesignator", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_workstation", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_automanual", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_flag", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("o_err_msg", System.Type.GetType("System.String")))

                Dim i As Integer
                Dim myDR As DataRow

                For i = 0 To dtModel.Rows.Count - 1
                    myDR = dtDJ.NewRow
                    myDR("p_dj_name") = dtModel.Rows(i)("DJ").ToString
                    dtDJ.Rows.Add(myDR)
                Next


                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.expand_bom_bydj"
                oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32).Value = LoginData.OrgID
                oda.InsertCommand.Parameters.Add("p_dj_name", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_identity", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                oda.InsertCommand.Parameters("p_dj_name").SourceColumn = "p_dj_name"
                oda.InsertCommand.Parameters("o_identity").SourceColumn = "o_identity"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_err_msg").SourceColumn = "o_err_msg"

                oda.InsertCommand.Connection.Open()
                oda.Update(dtDJ)                                        'DJ Data
                oda.InsertCommand.Connection.Close()

                Dim ErrMsg As String
                Dim DR() As DataRow = Nothing
                DR = dtDJ.Select(" o_flag = 'N' or o_flag = ' ' or o_flag IS Null ")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        If DR(i)("o_err_msg").ToString <> "" Then
                            ErrMsg = "Read DJ: " & DR(i)("p_dj_name") & "  error: " & DR(i)("o_err_msg").ToString
                            ErrorLogging("ProdPicking-DJReadProcessLabel", LoginData.User.ToUpper, ErrMsg, "I")
                        End If
                        DR(i).Delete()
                    Next
                    dtDJ.AcceptChanges()
                    ErrorMsg = ErrorMsg & ErrMsg

                    If dtDJ.Rows.Count = 0 Then Return ErrorMsg
                End If

                For i = 0 To dtDJ.Rows.Count - 1
                    dtDJ.Rows(i)("o_flag") = ""
                    dtDJ.Rows(i)("o_err_msg") = ""
                    dtDJ.Rows(i).AcceptChanges()
                    dtDJ.Rows(i).SetAdded()
                Next

                'This purpose is to delete temporary table data after reading data compeleted
                Dim dtDJData As New DataTable
                dtDJData = dtDJ.Copy

                'Collect PN for Data reading
                For i = 0 To dtMatl.Rows.Count - 1
                    DR = dtDJ.Select("p_dj_name = '" & dtMatl.Rows(i)("DJ").ToString & "'")
                    If DR.Length > 0 Then
                        myDR = dtItem.NewRow
                        myDR("p_identity") = DR(0)("o_identity").ToString
                        myDR("p_dj_name") = dtMatl.Rows(i)("DJ").ToString
                        myDR("p_com_item") = dtMatl.Rows(i)("MaterialNo").ToString
                        dtItem.Rows.Add(myDR)
                    End If
                Next

                oda.InsertCommand.Parameters.Clear()
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.get_process_label"
                oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32).Value = LoginData.OrgID
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_dj_name", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_com_item", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_floor", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_dj_start_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_dj_model", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_assy", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_req_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_usage_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_forming_code", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_refdesignator", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_workstation", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_automanual", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                oda.InsertCommand.Parameters("p_identity").SourceColumn = "p_identity"
                oda.InsertCommand.Parameters("p_dj_name").SourceColumn = "p_dj_name"
                oda.InsertCommand.Parameters("p_com_item").SourceColumn = "p_com_item"
                oda.InsertCommand.Parameters("o_floor").SourceColumn = "o_floor"
                oda.InsertCommand.Parameters("o_dj_start_qty").SourceColumn = "o_dj_start_qty"
                oda.InsertCommand.Parameters("o_dj_model").SourceColumn = "o_dj_model"
                oda.InsertCommand.Parameters("o_assy").SourceColumn = "o_assy"
                oda.InsertCommand.Parameters("o_req_qty").SourceColumn = "o_req_qty"
                oda.InsertCommand.Parameters("o_usage_qty").SourceColumn = "o_usage_qty"
                oda.InsertCommand.Parameters("o_forming_code").SourceColumn = "o_forming_code"
                oda.InsertCommand.Parameters("o_refdesignator").SourceColumn = "o_refdesignator"
                oda.InsertCommand.Parameters("o_workstation").SourceColumn = "o_workstation"
                oda.InsertCommand.Parameters("o_automanual").SourceColumn = "o_automanual"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_err_msg").SourceColumn = "o_err_msg"

                oda.InsertCommand.Connection.Open()
                oda.Update(dtItem)                                    'PN Data
                oda.InsertCommand.Connection.Close()

                myMOData.Items.Tables.Add(dtItem)

                DeleteBOMExpandData(dtDJData, LoginData.User)

            Catch ex As Exception
                ErrorLogging("ProdPicking-DJReadProcessLabel", LoginData.User, ErrorMsg & ex.Message & ex.Source, "E")
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function DeleteBOMExpandData(ByVal dtDJData As DataTable, ByVal User As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Insert()
            Dim ErrorMsg As String = "Delete BOM Data with error for ProcessLabel; "

            Try
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_invcc_pkg.delete_xxetr_bom_expand"
                oda.InsertCommand.Parameters.Add("p_identity", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                oda.InsertCommand.Parameters("p_identity").SourceColumn = "o_identity"
                oda.InsertCommand.Parameters("o_flag").SourceColumn = "o_flag"
                oda.InsertCommand.Parameters("o_err_msg").SourceColumn = "o_err_msg"

                oda.InsertCommand.Connection.Open()
                oda.Update(dtDJData)                                           'DJ Data
                oda.InsertCommand.Connection.Close()
                Return True

            Catch ex As Exception
                ErrorLogging("ProdPicking-DeleteBOMExpandData", User, ErrorMsg & ex.Message & ex.Source, "E")
                Return False
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function SaveProcessLabel(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim myDR As DataRow
            Dim dtLabel As New DataTable
            Dim ErrorMsg As String = "Save error for ProcessLabel; "

            dtLabel.Columns.Add(New Data.DataColumn("ProcessID", System.Type.GetType("System.String")))
            dtLabel.Columns.Add(New Data.DataColumn("CLID", System.Type.GetType("System.String")))
            dtLabel.Columns.Add(New Data.DataColumn("Content", System.Type.GetType("System.String")))

            Try
                Dim i, j As Integer
                Dim DR() As DataRow

                Dim ds As New DataSet
                Dim dtDJ As New DataTable
                Dim dtCLID As New DataTable

                Dim dtCode As New DataTable
                dtCode.Columns.Add(New Data.DataColumn("Code", System.Type.GetType("System.String")))

                dtDJ = myMOData.Items.Tables(0).Copy
                dtCLID = myMOData.Items.Tables(1).Copy
                myMOData.Items = New DataSet

                Dim LabelData As ProcessLabel
                Dim ProcessLabelID As String = ""

                For i = 0 To dtCLID.Rows.Count - 1
                    Dim DJ As String = dtCLID.Rows(i)("DJ").ToString
                    Dim CLID As String = dtCLID.Rows(i)("CLID").ToString
                    Dim MaterialNo As String = dtCLID.Rows(i)("MaterialNo").ToString
                    If DJ = "" Then Continue For

                    DR = Nothing
                    DR = dtDJ.Select("o_flag = 'Y' and p_dj_name='" & DJ & "' and p_com_item ='" & MaterialNo & "'")
                    If DR.Length = 0 Then Continue For

                    LabelData.MaterialNo = MaterialNo
                    LabelData.DJSet = DR(0)("o_floor") & " / " & DJ & " / " & DR(0)("o_dj_start_qty").ToString
                    LabelData.Model = DR(0)("o_dj_model").ToString

                    'Get Assy PN to Combine Model/Assy, but all no more than 56 characters     --12/19/2017
                    If DR(0)("o_assy").ToString <> "" Then
                        Dim AssyPN As String = ""
                        Dim myArry() As String
                        myArry = Split(DR(0)("o_assy").ToString, ",")
                        For j = 0 To myArry.Length - 1
                            'Assy: only extract the first 10 characters: XXX-XXXXXX
                            'Dim tmpPN As String = myArry(j).ToString.Trim
                            Dim tmpPN As String = Left(myArry(j).ToString.Trim, 10)
                            If tmpPN = "" Then Continue For

                            If AssyPN = "" Then
                                AssyPN = tmpPN
                            Else
                                Dim temp As String = AssyPN & "," & tmpPN
                                If temp.Length <= 56 Then
                                    AssyPN = AssyPN & "," & tmpPN
                                Else
                                    AssyPN = AssyPN & "..."
                                    Exit For
                                End If
                            End If
                        Next
                        LabelData.Model = LabelData.Model & " / " & AssyPN
                    End If

                    LabelData.ReqQty = DR(0)("o_req_qty").ToString
                    LabelData.UsageQty = DR(0)("o_usage_qty").ToString
                    'LabelData.FormingCode = DR(0)("o_forming_code").ToString
                    LabelData.RefDesignator = DR(0)("o_refdesignator").ToString
                    LabelData.WorkStation = DR(0)("o_workstation").ToString
                    LabelData.AutoManual = DR(0)("o_automanual").ToString

                    Dim PRArry() As String
                    Dim ProcessCode As String = DR(0)("o_forming_code").ToString
                    PRArry = Split(ProcessCode, "/")

                    'Clear the data before insert, this is very importment here
                    dtCode.Clear()

                    'ProcessCode, 如果FM开头：则取从开头取至第二个“-”后的4位; 
                    '如果不是FM开头：则取从第一个“-”后开始取, 取至第二个“-”后的4位
                    Dim StrCode As String = ""
                    For j = 0 To PRArry.Length - 1
                        Dim PRCode As String = PRArry(j).ToString.Trim
                        If PRCode = "" Then Continue For

                        Dim myCode As String
                        Dim myArry() As String
                        myArry = Split(PRCode, "-")

                        'FM01-CP0000-043R/FM02-CP0000-0458/HTSK1-CPBP00-112R/HTSK1-CPBP00-110
                        Dim FsCode As String = myArry(0).ToString
                        If Left(FsCode, 2) = "FM" Then
                            '如果是FM开头：则取从开头取至第二个“-”后的4位; 
                            If myArry.Length <= 2 Then
                                myCode = PRCode
                            Else
                                Dim StrIndex As Integer = myArry(0).Length + myArry(1).Length + 2
                                If PRCode.Length >= StrIndex + 4 Then
                                    myCode = Left(PRCode, StrIndex + 4)
                                Else
                                    myCode = PRCode
                                End If
                            End If

                        Else
                            '如果不是FM开头：则取从第一个“-”后开始取, 取至第二个“-”后的4位
                            If myArry.Length <= 2 Then
                                myCode = PRCode
                            Else
                                Dim StrIndex As Integer = myArry(0).Length + myArry(1).Length + 2
                                Dim FstIndex As Integer = myArry(0).Length + 2
                                Dim SecIndex As Integer = myArry(1).Length + 5
                                If PRCode.Length >= StrIndex + 4 Then
                                    myCode = Mid(PRCode, FstIndex, SecIndex)
                                Else
                                    myCode = Mid(PRCode, FstIndex)
                                End If
                            End If
                        End If

                        Dim dr1() As DataRow = Nothing
                        dr1 = dtCode.Select("Code ='" & myCode & "'")
                        If dr1.Length = 0 Then
                            myDR = dtCode.NewRow
                            myDR("Code") = myCode
                            dtCode.Rows.Add(myDR)
                            If StrCode = "" Then
                                StrCode = myCode
                            Else
                                'Get ProcessCode to Combine it, but no more than 67 characters     --12/19/2017
                                Dim temp As String = StrCode & "," & myCode
                                If temp.Length <= 67 Then
                                    StrCode = StrCode & "," & myCode
                                Else
                                    StrCode = StrCode & "..."
                                    Exit For
                                End If
                            End If

                        End If
                    Next
                    LabelData.FormingCode = StrCode


                    Dim tmpContent As String = ""
                    tmpContent = "^DJSet^" & LabelData.DJSet & "^Model^" & LabelData.Model & "^MaterialNo^" & LabelData.MaterialNo & "^ReqQty^" & LabelData.ReqQty & "^UsageQty^" & LabelData.UsageQty _
                                         & "^FormingCode^" & LabelData.FormingCode & "^RefDesignator^" & LabelData.RefDesignator & "^WorkStation^" & LabelData.WorkStation & "^AutoManual^" & LabelData.AutoManual

                    'Filter Special Characters for LabelData
                    tmpContent = FilterSpecial(tmpContent)

                    Dim Sqlstr As String
                    Dim BoxID As String = dtCLID.Rows(i)("BoxID").ToString
                    If BoxID <> "" Then                  'This is BoxID, then need to read all related CLIDs
                        Sqlstr = String.Format("Select CLID, BoxID from T_CLMaster with (nolock) where BoxID = '{0}' and OrgCode='{1}' and (SLOC<>'' or not (SLOC IS NULL)) ", BoxID, LoginData.OrgCode)
                        ds = da.ExecuteDataSet(Sqlstr, "CLIDs")

                        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                        ElseIf ds.Tables(0).Rows.Count > 0 Then
                            For j = 0 To ds.Tables(0).Rows.Count - 1
                                CLID = ds.Tables(0).Rows(j)("CLID").ToString

                                ProcessLabelID = GetTranHeaderID(LoginData, "PLID")
                                LabelData.ProcessLabelID = ProcessLabelID
                                If ProcessLabelID = "" Then Continue For

                                Dim LblContent As String = "ProcessLabelID^" & LabelData.ProcessLabelID & tmpContent
                                Sqlstr = String.Format("INSERT INTO T_RTSlip (TransactionID,OrgCode,RTNo,MaterialNo,PONo,RTContent,CreatedOn,CreatedBy) values ('{0}','{1}','{2}','{3}','{4}','{5}', getDate(),'{6}')", _
                                               ProcessLabelID, LoginData.OrgCode, CLID, MaterialNo, DJ, LblContent, LoginData.User.ToUpper)
                                da.ExecuteNonQuery(Sqlstr)

                                'Save Process Label Data for later printing use
                                myDR = dtLabel.NewRow
                                myDR("CLID") = CLID
                                myDR("ProcessID") = ProcessLabelID
                                myDR("Content") = LblContent
                                dtLabel.Rows.Add(myDR)
                            Next
                        End If
                    Else
                        ProcessLabelID = GetTranHeaderID(LoginData, "PLID")
                        LabelData.ProcessLabelID = ProcessLabelID
                        If ProcessLabelID = "" Then Continue For

                        Dim LblContent As String = "ProcessLabelID^" & LabelData.ProcessLabelID & tmpContent
                        Sqlstr = String.Format("INSERT INTO T_RTSlip (TransactionID,OrgCode,RTNo,MaterialNo,PONo,RTContent,CreatedOn,CreatedBy) values ('{0}','{1}','{2}','{3}','{4}','{5}', getDate(),'{6}')", _
                                       ProcessLabelID, LoginData.OrgCode, CLID, MaterialNo, DJ, LblContent, LoginData.User.ToUpper)
                        da.ExecuteNonQuery(Sqlstr)

                        'Save Process Label Data for later printing use
                        myDR = dtLabel.NewRow
                        myDR("CLID") = CLID
                        myDR("ProcessID") = ProcessLabelID
                        myDR("Content") = LblContent
                        dtLabel.Rows.Add(myDR)
                    End If
                Next

                myMOData.Items.Tables.Add(dtLabel)
                SaveProcessLabel = True

            Catch ex As Exception
                ErrorLogging("ProdPicking-SaveProcessLabel", LoginData.User, ErrorMsg & ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using

    End Function

    Public Function PrintProcessLabel(ByVal LoginData As ERPLogin, ByVal ds As DataSet, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            PrintProcessLabel = False

            Try
                Dim i, aa As Integer
                Dim arryFile() As String
                arryFile = Split(ProcessLabelFile, "\")
                Dim LabelFile As String = arryFile(UBound(arryFile))

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Dim strContent As String = ds.Tables(0).Rows(i)("Content").ToString

                    'Filter Special Characters for Label Data
                    strContent = FilterSpecial(strContent)

                    Dim sqlstr As String
                    sqlstr = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                    da.ExecuteScalar(sqlstr)
                    Sleep(5)
                Next
                PrintProcessLabel = True

            Catch ex As Exception
                ErrorLogging("ProdPicking-PrintProcessLabel", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                PrintProcessLabel = False
            End Try
        End Using

    End Function

    Public Function ReadProcessLabel(ByVal LoginData As ERPLogin, ByVal LabelID As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec dbo.sp_ReadProcessLabel '{0}','{1}'", LoginData.OrgCode, LabelID)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"dtMsg", "dtLabel"}
                ds = da.ExecuteDataSet(sql, tables)

                'Process Label is valid
                If ds.Tables(0).Rows(0)(0).ToString = "Y" Then
                    Dim strContent As String = ds.Tables(1).Rows(0)("Content").ToString
                    'ProcessLabelID^PL00000002^DJSet^TIS300 / P2017834 / 4368^Model^700-013834-0000 / ^MaterialNo^621-003482-0103
                    '^ReqQty^4368^UsageQty^1^FormingCode^^RefDesignator^INSE-TRAY1-1^WorkStation^^AutoManual^A

                    Dim arryFile() As String
                    arryFile = Split(strContent, "^")

                    ds.Tables(1).Rows(0)("DJSet") = arryFile(3).ToString
                    ds.Tables(1).Rows(0)("Model") = arryFile(5).ToString
                    ds.Tables(1).Rows(0)("MaterialNo") = arryFile(7).ToString
                    ds.Tables(1).Rows(0)("FormingCode") = arryFile(13).ToString
                    ds.Tables(1).AcceptChanges()
                End If
                Return ds

            Catch ex As Exception
                ErrorLogging("ProdPicking-ReadProcessLabel", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function PCBSlotLightOn(ByVal LoginData As ERPLogin, ByRef myMOData As MOData) As String
        Using da As DataAccess = GetDataAccess()

            Try
                PCBSlotLightOn = ""
                myMOData.PickFlag = "N"                                               'Light On Slot has Error, so no need to Light Off

                Dim ds As New DataSet
                ds = myMOData.Items.Copy
                myMOData.Items = New DataSet

                Dim i As Integer
                Dim DR() As DataRow
                Dim sda As SqlClient.SqlDataAdapter = da.Sda_Insert()

                Try
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        ds.Tables(0).Rows(i).SetAdded()
                    Next

                    sda.InsertCommand.CommandType = CommandType.StoredProcedure
                    sda.InsertCommand.CommandText = "sp_ReadSlotForPCB"
                    sda.InsertCommand.CommandTimeout = TimeOut_M5

                    sda.InsertCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50).Value = LoginData.OrgCode
                    sda.InsertCommand.Parameters.Add("@SubInv", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@Locator", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@Slot", SqlDbType.VarChar, 50)

                    sda.InsertCommand.Parameters("@SubInv").SourceColumn = "SubInv"
                    sda.InsertCommand.Parameters("@Locator").SourceColumn = "Locator"
                    sda.InsertCommand.Parameters("@Slot").SourceColumn = "Slot"
                    sda.InsertCommand.Parameters("@Slot").Direction = ParameterDirection.Output

                    sda.InsertCommand.Connection.Open()
                    sda.Update(ds.Tables(0))
                    sda.InsertCommand.Connection.Close()

                    DR = ds.Tables(0).Select("Slot <> '' ")
                    If DR.Length = 0 Then PCBSlotLightOn = "No available Slot need to Light On."

                Catch ex As Exception
                    ErrorLogging("ProdPicking-call-sp_ReadSlotForPCB", LoginData.User, ex.Message & ex.Source, "E")
                    PCBSlotLightOn = "Read Slot for PCB error"
                Finally
                    If sda.InsertCommand.Connection.State <> ConnectionState.Closed Then sda.InsertCommand.Connection.Close()
                End Try

                'Return if there has error
                If PCBSlotLightOn <> "" Then Return PCBSlotLightOn

                Dim dsSlot As New DataSet("DS")
                Dim myDR As DataRow
                Dim dtSlot As DataTable
                dtSlot = New DataTable("dtSlot")
                dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                dsSlot.Tables.Add(dtSlot)

                For i = 0 To DR.Length - 1
                    myDR = dtSlot.NewRow()
                    myDR("slot") = DR(i)("Slot").ToString
                    dtSlot.Rows.Add(myDR)
                Next

                '==Code--0/1/2(Off/On/blink. data type: integer)
                '==Interval--time(data type: integer. 0:not limit)
                Dim Code As Integer = 1
                Dim myWMS As WMS = New WMS
                myWMS.LEDControlBySlot(dsSlot, Code, 0)

                PCBSlotLightOn = "Y"
                myMOData.PickFlag = "Y"                    'Light On Slot already, so need to Light Off after CLID Confirm

                myMOData.Items.Merge(ds)

            Catch ex As Exception
                ErrorLogging("ProdPicking-PCBSlotLightOn", LoginData.User, ex.Message & ex.Source, "E")
                PCBSlotLightOn = "Light On Slot for PCB error"
            End Try

        End Using
    End Function

#End Region

#Region "SMTMOConfirm"
    Public Function CheckUTurnFlag(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim v_count As Integer
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("select count(1) v_count from T_InvEventHeader where EventID = '{0}' and DestSubInv in ('ZTU RW1','ZTU RW2') ", mySMOData.EventID)
                v_count = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                If v_count > 0 Then
                    CheckUTurnFlag = True
                Else
                    CheckUTurnFlag = False
                End If
            Catch ex As Exception
                ErrorLogging("WMSMO-CheckUTurnFlag", mySMOData.User, "EventID:" & mySMOData.EventID & "," & ex.ToString, "E")
                mySMOData.RtnMsg = "Check U-Turn Flag with error."
                Return False
            End Try
        End Using
    End Function
    Public Function CheckBoxID(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim v_count As Integer
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("select count(1) v_count from T_CLMaster_Packing where BoxID = '{0}' and OrgCode = '{1}' ", mySMOData.BoxID, mySMOData.OrgCode)
                v_count = Convert.ToInt32(da.ExecuteScalar(Sqlstr))
                If v_count > 0 Then
                    CheckBoxID = True
                Else
                    CheckBoxID = False
                End If
            Catch ex As Exception
                ErrorLogging("WMSMO-CheckBoxID", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Check BoxID with error."
                Return False
            End Try
        End Using
    End Function
    Public Function MissingPickExactCLID(ByVal EventID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            sda = da.Sda_Sele()
            Dim spResult As String
            Try
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_MissingPickExactCLID"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@eventid", SqlDbType.VarChar, 250).Value = EventID
                sda.SelectCommand.Parameters.Add("@nextstep", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                spResult = sda.SelectCommand.Parameters("@nextstep").Value.ToString()
                Return spResult
            Catch oe As Exception
                ErrorLogging("CheckCLIDExpFormat", "", oe.Message & oe.Source, "E")
                Return "E"
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function CheckProdLine(ByRef mySMOData As SMTData) As String
        Using da As DataAccess = GetDataAccess()
            CheckProdLine = ""

            Dim Sqlstr As String
            Dim ds As New DataSet

            Try

                Sqlstr = String.Format("exec sp_WMSCheckDock '{0}', '{1}', N'{2}', '{3}' ", mySMOData.UserProdLine, mySMOData.OrgCode, mySMOData.User, mySMOData.CheckDock)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Msg", "Header", "Item"}
                ds = da.ExecuteDataSet(sql, tables)

                CheckProdLine = ds.Tables(0).Rows(0)(0).ToString
                mySMOData.RtnMsg = ds.Tables(0).Rows(0)(1).ToString

                mySMOData.dsItem = New DataSet
                mySMOData.dsItem.Merge(ds.Tables(1))
                mySMOData.dsItem.Merge(ds.Tables(2))

            Catch ex As Exception
                ErrorLogging("WMSMO-CheckProdLine", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Check ProdLine with error."
                Return "N"
            End Try
        End Using

    End Function

    Public Function ConfirmCLID(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ConfirmCLID = False

            Dim ds As New DataSet

            Try
                Dim Sqlstr, myResult As String
                Sqlstr = String.Format("exec sp_WMSReadEventCLID '{0}', '{1}' ", mySMOData.EventID, mySMOData.CLID)
                ds = da.ExecuteDataSet(Sqlstr, "CLID")
                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    mySMOData.RtnMsg = "Coundn't find CLID " & mySMOData.CLID & " in T_CLMaster"
                    Return False
                End If

                'For EventType = "SO Shipment", need to check if the Raw Item has COO or not
                If mySMOData.EventType = "SO Shipment" Then
                    Dim COO As String = ds.Tables(0).Rows(0)("COO").ToString
                    Dim Matl As String = ds.Tables(0).Rows(0)("Item").ToString
                    Dim ItemType As String = ds.Tables(0).Rows(0)("ItemType").ToString

                    Dim ErrMsg As String
                    If ItemType = "RM" AndAlso COO = "" Then
                        ErrMsg = "^WMS-37@ " & Matl & " ^WMS-38@"
                        mySMOData.RtnMsg = ErrMsg
                        Return False
                    End If
                End If

                'If Event Header Status is 'LightOn', then Light Off the LED
                Dim InvSlot As String = ds.Tables(0).Rows(0)("InvSlot").ToString
                Dim EvtStatus As String = ds.Tables(0).Rows(0)("EvtStatus").ToString

                Dim LightOn As Boolean = False
                If EvtStatus = "LightOn" AndAlso InvSlot <> "" Then
                    'SlotLightOn(InvSlot, LightOn, mySMOData.User)

                    Dim dsSlot As New DataSet("DS")
                    Dim myDR As DataRow
                    Dim dtSlot As DataTable
                    dtSlot = New DataTable("dtSlot")
                    dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                    dsSlot.Tables.Add(dtSlot)

                    myDR = dtSlot.NewRow()
                    myDR("slot") = InvSlot
                    dtSlot.Rows.Add(myDR)

                    '==Code--0/1/2(Off/On/blink. data type: integer)
                    '==Interval--time(data type: integer. 0:not limit)
                    Dim Code As Integer = 0
                    Dim myWMS As WMS = New WMS
                    myWMS.LEDControlBySlot(dsSlot, Code, 5)
                End If


                Sqlstr = String.Format("exec sp_WMSConfirmCLID '{0}', '{1}', '{2}', '{3}', N'{4}', '{5}' ", mySMOData.OrgCode, mySMOData.EventID, mySMOData.CLID, mySMOData.ActionType, mySMOData.User, mySMOData.BoxID)
                myResult = da.ExecuteScalar(Sqlstr).ToString

                If Left(myResult, 2) <> "Y/" Then
                    mySMOData.RtnMsg = myResult
                    Return False
                End If

                ConfirmCLID = True
                Dim CLIDLists As String = Microsoft.VisualBasic.Mid(myResult, 3)
                mySMOData.RtnMsg = CLIDLists

                'No need to Light on the LED as there is no light here right now   -- 10/4/2017
                ''Light on the LED for the Slot if this CLID is Primary and it was not Splitted
                'Dim DockSlot As String
                'DockSlot = ds.Tables(0).Rows(0)("DockSlot").ToString
                'If DockSlot <> "" AndAlso CLIDLists = "" Then
                '    LightOn = True
                '    Dim Interval As Integer = 8                                              'User required to LightOn 8 seconds for Docking
                '    'SlotLightOn(DockSlot, LightOn, mySMOData.User)
                '    DockSlotLightOn(DockSlot, LightOn, mySMOData.User, Interval)
                'End If

            Catch ex As Exception
                ErrorLogging("WMSMO-ConfirmCLID", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Confirm CLID with error."
                Return False
            End Try
        End Using

    End Function

    'Public Function ConfirmMO(ByVal LoginData As ERPLogin, ByRef mySMOData As SMTData) As Boolean
    '    Using da As DataAccess = GetDataAccess()

    '        ConfirmMO = False
    '        mySMOData.RtnMsg = ""

    '        Dim Msg As String = "MO: " & mySMOData.MONo & "; "
    '        Dim ErrMsg As String = "Confirm MO with error for MO " & mySMOData.MONo

    '        Try
    '            Dim dtCLIDs As DataTable = New DataTable
    '            dtCLIDs = mySMOData.dsItem.Tables(0)

    '            Dim ds As DataSet = New DataSet
    '            ds.Tables.Add("MatlList")
    '            ds.Tables(0).Columns.Add(New Data.DataColumn("MO", System.Type.GetType("System.String")))
    '            ds.Tables(0).Columns.Add(New Data.DataColumn("Item", System.Type.GetType("System.String")))
    '            ds.Tables(0).Columns.Add(New Data.DataColumn("Qty", System.Type.GetType("System.Decimal")))
    '            ds.Tables(0).Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
    '            ds.Tables(0).Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))

    '            Dim i As Integer
    '            Dim myDR As DataRow
    '            For i = 0 To dtCLIDs.Rows.Count - 1
    '                Dim DR() As DataRow
    '                DR = ds.Tables(0).Select("Item = '" & dtCLIDs.Rows(i)("Item") & "'")
    '                If DR.Length = 0 Then
    '                    myDR = ds.Tables(0).NewRow
    '                    myDR("MO") = mySMOData.MONo
    '                    myDR("Item") = dtCLIDs.Rows(i)("Item")
    '                    myDR("Status") = ""
    '                    myDR("Message") = ""

    '                    Dim Qty As Decimal = dtCLIDs.Rows(i)("Qty")
    '                    myDR("Qty") = Math.Round(Qty, 6)
    '                    ds.Tables(0).Rows.Add(myDR)
    '                Else
    '                    Dim TotalQty As Decimal = DR(0)("Qty") + dtCLIDs.Rows(i)("Qty")
    '                    DR(0)("Qty") = Math.Round(TotalQty, 6)
    '                    DR(0).AcceptChanges()
    '                    DR(0).SetAdded()
    '                End If
    '            Next

    '            Dim SaveMO As String = ""
    '            Try
    '                SaveMO = ProcessSMTMO(LoginData, ds)
    '                'SaveMO = TransactSMTMO(LoginData, ds)
    '            Catch ex As Exception
    '                ErrorLogging("WMSMO-ConfirmMO1", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '                SaveMO = ErrMsg
    '            End Try
    '            If SaveMO <> "Y" Then
    '                mySMOData.RtnMsg = SaveMO
    '                Return False
    '            End If


    '            Dim ActionType As String = "Confirm"           'Confirm the Whole MO and update data
    '            Try
    '                Dim Sqlstr As String
    '                Sqlstr = String.Format("exec sp_WMSConfirmMO '{0}', '{1}', '{2}', N'{3}' ", mySMOData.OrgCode, mySMOData.EventID, ActionType, mySMOData.User)
    '                SaveMO = da.ExecuteScalar(Sqlstr).ToString
    '            Catch ex As Exception
    '                ErrorLogging("WMSMO-sp_WMSConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '                SaveMO = ErrMsg
    '            End Try
    '            If SaveMO <> "Y" Then
    '                mySMOData.RtnMsg = SaveMO
    '                Return False
    '            End If

    '            ConfirmMO = True

    '        Catch ex As Exception
    '            ErrorLogging("WMSMO-ConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
    '            mySMOData.RtnMsg = ErrMsg
    '            Return False
    '        End Try
    '    End Using

    'End Function
    Public Function ConfirmMO(ByVal LoginData As ERPLogin, ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()

            ConfirmMO = False
            mySMOData.RtnMsg = ""

            Dim Msg As String = "MO: " & mySMOData.MONo & "; "
            Dim ErrMsg As String = "Confirm MO with error for MO " & mySMOData.MONo

            Try
                Dim dsItem As New DataSet
                Dim ActionType As String = "Confirm"           'Confirm the Whole MO and update data

                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_WMSReadEventID '{0}', '{1}' ", mySMOData.EventID, ActionType)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"MatlList", "MissPN"}
                dsItem = da.ExecuteDataSet(sql, tables)
                If dsItem Is Nothing OrElse dsItem.Tables.Count = 0 Then
                    mySMOData.RtnMsg = "Invalid EventID " & mySMOData.EventID
                    Return False
                End If

                Dim MissingFlag As Boolean = False
                If dsItem.Tables(1).Rows.Count > 0 Then MissingFlag = True

                Dim dtCLIDs As New DataTable
                Dim dtMItem As New DataTable
                dtCLIDs = dsItem.Tables(0).Copy
                dtMItem = dsItem.Tables(1).Copy

                'First Process Normal Items without Missing
                If dtCLIDs.Rows.Count > 0 Then
                    mySMOData.ActionType = "1"                             'Normal Items without Missing
                    mySMOData.dsItem = New DataSet
                    mySMOData.dsItem.Merge(dtCLIDs)
                    If CheckUTurnFlag(mySMOData) = True Then
                        mySMOData.ActionType = "2"
                        ConfirmMO = SaveSMTMOData(LoginData, mySMOData, True)
                    Else
                    ConfirmMO = SaveSMTMOData(LoginData, mySMOData, MissingFlag)
                    End If
                    If ConfirmMO = False Then Return False
                    If mySMOData.CheckDock = "R" Then Exit Function
                End If


                'Second Process those Items with Missing 
                If MissingFlag = True Then
                    mySMOData.ActionType = "2"                             'Items with Missing
                    dtMItem.TableName = "MatlList"
                    mySMOData.dsItem = New DataSet
                    mySMOData.dsItem.Merge(dtMItem)

                    ConfirmMO = SaveSMTMOData(LoginData, mySMOData, MissingFlag)
                    If ConfirmMO = False Then Return False
                End If


                'If All Items are Missing, then we just Close the Event Header
                If dtCLIDs.Rows.Count = 0 AndAlso MissingFlag = False Then
                    Dim myResult As String = ""
                    Try
                        Dim ConfirmType As String = mySMOData.EventType
                        Sqlstr = String.Format("exec sp_WMSConfirmMO '{0}', '{1}', '{2}', N'{3}' ", mySMOData.OrgCode, mySMOData.EventID, ConfirmType, mySMOData.User)
                        myResult = da.ExecuteScalar(Sqlstr).ToString
                    Catch ex As Exception
                        ErrorLogging("WMSMO-ConfirmMO-sp_WMSConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        myResult = ErrMsg
                    End Try
                    If myResult <> "Y" Then
                        mySMOData.RtnMsg = myResult
                        Return False
                    End If
                End If

                ConfirmMO = True

            Catch ex As Exception
                ErrorLogging("WMSMO-ConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

    Public Function ReConfirmMO(ByVal LoginData As ERPLogin, ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()

            ReConfirmMO = False
            mySMOData.RtnMsg = ""

            Dim Msg As String = "MO: " & mySMOData.MONo & "; "
            Dim ErrMsg As String = "ReConfirm MO with error for MO " & mySMOData.MONo

            Try
                Dim dsItem As New DataSet
                Dim ActionType As String = "Confirm"           'Confirm the Whole MO and update data

                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_WMSReadEventID '{0}', '{1}' ", mySMOData.EventID, ActionType)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"MatlList", "MissPN"}
                dsItem = da.ExecuteDataSet(sql, tables)
                If dsItem Is Nothing OrElse dsItem.Tables.Count = 0 Then
                    mySMOData.RtnMsg = "Invalid EventID " & mySMOData.EventID
                    Return False
                End If

                Dim MissingFlag As Boolean = False
                If dsItem.Tables(1).Rows.Count > 0 Then MissingFlag = True

                Dim dtCLIDs As New DataTable
                Dim dtMItem As New DataTable
                dtCLIDs = dsItem.Tables(0).Copy
                dtMItem = dsItem.Tables(1).Copy

                'First Process Normal Items without Missing
                If dtCLIDs.Rows.Count > 0 Then
                    mySMOData.ActionType = "1"                             'Normal Items without Missing
                    mySMOData.dsItem = New DataSet
                    mySMOData.dsItem.Merge(dtCLIDs)
                    ReConfirmMO = SaveSMTMOData(LoginData, mySMOData, MissingFlag)
                    If ReConfirmMO = False Then Return False
                End If


                'Second Process those Items with Missing 
                If MissingFlag = True Then
                    mySMOData.ActionType = "2"                             'Items with Missing
                    dtMItem.TableName = "MatlList"
                    mySMOData.dsItem = New DataSet
                    mySMOData.dsItem.Merge(dtMItem)

                    ReConfirmMO = SaveSMTMOData(LoginData, mySMOData, MissingFlag)
                    If ReConfirmMO = False Then Return False
                End If


                'If All Items are Missing, then we just Close the Event Header
                If dtCLIDs.Rows.Count = 0 AndAlso MissingFlag = False Then
                    Dim myResult As String = ""
                    Try
                        Dim ConfirmType As String = mySMOData.EventType
                        Sqlstr = String.Format("exec sp_WMSConfirmMO '{0}', '{1}', '{2}', N'{3}' ", mySMOData.OrgCode, mySMOData.EventID, ConfirmType, mySMOData.User)
                        myResult = da.ExecuteScalar(Sqlstr).ToString
                    Catch ex As Exception
                        ErrorLogging("WMSMO-ReConfirmMO-sp_WMSConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        myResult = ErrMsg
                    End Try
                    If myResult <> "Y" Then
                        mySMOData.RtnMsg = myResult
                        Return False
                    End If
                End If

                ReConfirmMO = True

            Catch ex As Exception
                ErrorLogging("WMSMO-ReConfirmMO", LoginData.User, Msg & ex.Message & ex.Source, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

    Public Function SaveSMTMOData(ByVal LoginData As ERPLogin, ByRef mySMOData As SMTData, ByVal MissingFlag As Boolean) As Boolean
        Using da As DataAccess = GetDataAccess()

            SaveSMTMOData = False

            Dim Msg As String = "MO: " & mySMOData.MONo & "; "
            Dim ErrMsg As String = "Confirm MO with error for MO " & mySMOData.MONo

            Try
                Dim dtLot As DataTable
                dtLot = New Data.DataTable("Lots")
                dtLot.Columns.Add(New Data.DataColumn("Item", System.Type.GetType("System.String")))
                dtLot.Columns.Add(New Data.DataColumn("RTLot", System.Type.GetType("System.String")))

                Dim dtCLIDs As New DataTable
                Dim ActionType As String = mySMOData.EventType           'Confirm the Whole MO and update data
                dtCLIDs = mySMOData.dsItem.Tables(0).Copy

                Dim ds As New DataSet
                Dim dtMO As New DataTable
                dtMO = dtCLIDs.Clone
                ds.Tables.Add(dtMO)

                Dim i, j As Integer
                Dim myDR As DataRow
                Dim dtCOO As New DataTable
                dtCOO = dtCLIDs.Clone

                Dim Sqlstr As String
                Dim myAction As String = mySMOData.ActionType

                For i = 0 To dtCLIDs.Rows.Count - 1
                    Dim RTLot As String = dtCLIDs.Rows(i)("RTLot").ToString
                    Sqlstr = "Item = '" & dtCLIDs.Rows(i)("Item") & "' and RTLot = '" & RTLot & "'"

                    'Collect Items without Missing
                    Dim DR() As DataRow
                    If MissingFlag = True Then
                        DR = dtLot.Select(Sqlstr)
                        If DR.Length = 0 Then
                            myDR = dtLot.NewRow
                            myDR("Item") = dtCLIDs.Rows(i)("Item")
                            myDR("RTLot") = RTLot
                            dtLot.Rows.Add(myDR)
                        End If
                    End If

                    'Collect COO data for SOShipment with Raw Material
                    If myAction = "1" Then Sqlstr = "Item = '" & dtCLIDs.Rows(i)("Item") & "'"
                    DR = Nothing
                    DR = ds.Tables(0).Select(Sqlstr)
                    If DR.Length = 0 Then
                        myDR = ds.Tables(0).NewRow
                        myDR("MO") = mySMOData.MONo
                        myDR("DJ") = ""
                        If ActionType = "MO Confirm" Then
                            myDR("DJ") = mySMOData.EventID                         'Save EventID here only for MO Confirm
                        End If
                        myDR("Item") = dtCLIDs.Rows(i)("Item")
                        myDR("RTLot") = RTLot
                        If myAction = "1" Then myDR("RTLot") = ""
                        myDR("CLID") = ""
                        myDR("Status") = ""

                        Dim Qty As Decimal = dtCLIDs.Rows(i)("Qty")
                        myDR("Qty") = Math.Round(Qty, 6)
                        ds.Tables(0).Rows.Add(myDR)
                    Else
                        Dim TotalQty As Decimal = DR(0)("Qty") + dtCLIDs.Rows(i)("Qty")
                        DR(0)("Qty") = Math.Round(TotalQty, 6)
                        DR(0).AcceptChanges()
                        DR(0).SetAdded()
                    End If

                    'Collect COO data for SOShipment with Raw Material
                    If myAction = "1" AndAlso ActionType = "SO Shipment" Then
                        Dim ItemType As String = dtCLIDs.Rows(i)("ItemType").ToString
                        If ItemType = "RM" Then
                            dtCOO.ImportRow(dtCLIDs.Rows(i))
                        End If
                    End If
                Next

                Dim SaveMO As String = ""
                Dim myResult As String = ""

                If myAction = "1" Then
                    Dim dsMO As New DataSet
                    dsMO.Merge(ds)

                    Try
                        SaveMO = ProcessSMTMO(LoginData, ds, mySMOData)
                    Catch ex As Exception
                        ErrorLogging("WMSMO-SaveSMTMOData1", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        SaveMO = ErrMsg
                    End Try

                    'Flag R: need to send data to Oracle again only for MO Confirm
                    If SaveMO = "R" Then
                        Try
                            'SaveMO = ProcessSMTMO(LoginData, dsMO)
                            mySMOData.CheckDock = "R"                                                               'set Flag for later use
                            SaveSMTMOData = ReConfirmMO(LoginData, mySMOData)
                            Return SaveSMTMOData
                            Exit Function
                        Catch ex As Exception
                            ErrorLogging("WMSMO-ReConfirmMO2", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        End Try
                    End If

                    If SaveMO <> "Y" Then
                        mySMOData.RtnMsg = SaveMO
                        Return False
                    End If

                    'Insert data RM items / COO into Oracle customized table if there exists
                    If ActionType = "SO Shipment" Then
                        If dtCOO.Rows.Count > 0 Then
                            RecordSMTCOO(LoginData, dtCOO)
                        End If
                    End If

                    Dim sp As String = ""
                    Try
                        If MissingFlag = False Then
                            sp = "sp_WMSConfirmMO1"
                            Sqlstr = String.Format("exec sp_WMSConfirmMO '{0}', '{1}', '{2}', N'{3}' ", mySMOData.OrgCode, mySMOData.EventID, ActionType, mySMOData.User)
                        Else
                            Dim Items As New DataSet
                            Items.Tables.Add(dtLot)
                            Items.DataSetName = "Items"

                            sp = "sp_WMSConfirmMOByLot1"
                            Sqlstr = String.Format("exec sp_WMSConfirmMOByLot '{0}', '{1}', '{2}', N'{3}', N'{4}' ", mySMOData.OrgCode, mySMOData.EventID, ActionType, mySMOData.User, DStoXML(Items))
                        End If
                        myResult = da.ExecuteScalar(Sqlstr).ToString
                    Catch ex As Exception
                        ErrorLogging("WMSMO-" & sp, LoginData.User, Msg & ex.Message & ex.Source, "E")
                        myResult = ErrMsg
                    End Try
                    If myResult <> "Y" Then
                        mySMOData.RtnMsg = myResult
                        Return False
                    End If
                End If


                If myAction = "2" Then
                    Try
                        SaveMO = ProcessMissItem(LoginData, ds)
                    Catch ex As Exception
                        ErrorLogging("WMSMO-SaveSMTMOData2", LoginData.User, Msg & ex.Message & ex.Source, "E")
                        SaveMO = ErrMsg
                    End Try

                    Dim drLot() As DataRow
                    dtLot.Clear()
                    drLot = ds.Tables(0).Select("Status = 'Y' ")
                    If drLot.Length > 0 Then
                        For i = 0 To drLot.Length - 1
                            Dim DR() As DataRow
                            Dim Matl As String = drLot(i)("Item").ToString
                            Dim RTLot As String = drLot(i)("RTLot").ToString
                            DR = dtLot.Select("Item = '" & Matl & "' and RTLot = '" & RTLot & "'")
                            If DR.Length = 0 Then
                                myDR = dtLot.NewRow
                                myDR("Item") = Matl
                                myDR("RTLot") = RTLot
                                dtLot.Rows.Add(myDR)
                            End If

                            If ActionType = "SO Shipment" Then
                                DR = Nothing
                                DR = dtCLIDs.Select("Item = '" & Matl & "' and RTLot = '" & RTLot & "'")
                                If DR.Length > 0 Then
                                    For j = 0 To DR.Length - 1
                                        If DR(j)("ItemType") = "RM" Then
                                            dtCOO.ImportRow(DR(j))
                                        End If
                                    Next
                                End If
                            End If

                        Next

                        'Insert data RM items / COO into Oracle customized table if there exists
                        If ActionType = "SO Shipment" Then
                            If dtCOO.Rows.Count > 0 Then
                                RecordSMTCOO(LoginData, dtCOO)
                            End If
                        End If

                        Try
                            Dim Items As New DataSet
                            Items.Tables.Add(dtLot)
                            Items.DataSetName = "Items"
                            myResult = ""

                            Sqlstr = String.Format("exec sp_WMSConfirmMOByLot '{0}', '{1}', '{2}', N'{3}', N'{4}' ", mySMOData.OrgCode, mySMOData.EventID, ActionType, mySMOData.User, DStoXML(Items))
                            myResult = da.ExecuteScalar(Sqlstr).ToString
                        Catch ex As Exception
                            ErrorLogging("WMSMO-sp_WMSConfirmMOByLot2", LoginData.User, Msg & ex.Message & ex.Source, "E")
                            myResult = ErrMsg
                        End Try
                    End If
                    If SaveMO <> "Y" Then
                        mySMOData.RtnMsg = SaveMO
                        Return False
                    End If
                    If myResult <> "Y" Then
                        mySMOData.RtnMsg = myResult
                        Return False
                    End If
                End If

                SaveSMTMOData = True

            Catch ex As Exception
                ErrorLogging("WMSMO-SaveSMTMOData", LoginData.User, Msg & ex.Message & ex.Source, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

    'Public Function ProcessSMTMO(ByVal LoginData As ERPLogin, ByVal ds As DataSet) As String
    '    Using da As DataAccess = GetDataAccess()

    '        Try
    '            Dim MO As String = ds.Tables(0).Rows(0)("MO").ToString

    '            Dim oda As OracleDataAdapter
    '            Dim comm As OracleCommand

    '            Dim OrgID As String = GetOrgID(LoginData.OrgCode)

    '            Try
    '                oda = New OracleDataAdapter
    '                comm = da.Ora_Command_Trans()
    '            Catch ex As Exception
    '                ErrorLogging("ProcessSMTMO-OpenException", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
    '                ProcessSMTMO = ex.Message
    '                Return ProcessSMTMO
    '                Exit Function
    '            End Try

    '            Try
    '                comm.CommandType = CommandType.StoredProcedure
    '                comm.CommandText = "apps.XXETR_wip_pkg.initialize"
    '                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
    '                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
    '                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
    '                comm.ExecuteOracleNonQuery("")
    '                comm.Parameters.Clear()

    '                comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo"
    '                'comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo_by_lotnum"
    '                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
    '                comm.Parameters.Add("p_org_id", OracleType.VarChar, 20).Value = OrgID                        'LoginData.OrgCode
    '                comm.Parameters.Add("p_move_order", OracleType.VarChar, 50).Value = MO
    '                comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
    '                'comm.Parameters.Add("p_lot_num", OracleType.VarChar, 50)
    '                comm.Parameters.Add("p_quantity", OracleType.Double)
    '                comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
    '                comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

    '                comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
    '                comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

    '                comm.Parameters("p_item_num").SourceColumn = "Item"
    '                'comm.Parameters("p_lot_num").SourceColumn = "RTLot"
    '                comm.Parameters("p_quantity").SourceColumn = "Qty"

    '                comm.Parameters("o_success_flag").SourceColumn = "Status"
    '                comm.Parameters("o_error_mssg").SourceColumn = "Message"

    '                oda.InsertCommand = comm
    '                oda.Update(ds.Tables("MatlList"))

    '                If ds.Tables("MatlList").Rows.Count = 0 Then
    '                    comm.Transaction.Rollback()
    '                    comm.Connection.Close()
    '                    ProcessSMTMO = "No data is submitted in Oracle"
    '                    Return ProcessSMTMO
    '                End If

    '                Dim DR() As DataRow = Nothing
    '                DR = ds.Tables("MatlList").Select(" Status = 'N' or Status = ' ' or Status IS Null ")
    '                If DR.Length = 0 Then
    '                    comm.Transaction.Commit()
    '                    comm.Connection.Close()
    '                    ProcessSMTMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
    '                    Return ProcessSMTMO
    '                End If

    '                'Rollback and record error message if Oracle Data update failed
    '                comm.Transaction.Rollback()
    '                comm.Connection.Close()

    '                Dim i As Integer
    '                For i = 0 To DR.Length - 1
    '                    Dim ErrMsg As String = ""
    '                    ErrMsg = "MO " & MO & " for Item: " & DR(i)("Item").ToString & " and RTLot " & DR(i)("RTLot").ToString & " with IssuedQty " & DR(i)("Qty").ToString & " and flag: " & DR(i)("Status").ToString & " and error message; " & DR(i)("Message").ToString
    '                    ErrorLogging("WMSMO-ProcessSMTMO1", LoginData.User, ErrMsg, "I")
    '                Next

    '                ProcessSMTMO = DR(0)("Message").ToString
    '                If ProcessSMTMO = "" Then
    '                    ProcessSMTMO = "No data is submitted in Oracle"
    '                    Return ProcessSMTMO
    '                End If

    '                If ProcessSMTMO.ToUpper.Contains("ORA-00054") Then
    '                    ProcessSMTMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(0)("Item").ToString
    '                Else
    '                    ProcessSMTMO = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
    '                End If
    '                Return ProcessSMTMO                                'Return Error message

    '            Catch ex As Exception
    '                ErrorLogging("WMSMO-ProcessSMTMO", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
    '                comm.Transaction.Rollback()
    '                ProcessSMTMO = ex.Message
    '            Finally
    '                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
    '            End Try

    '        Catch ex As Exception
    '            ErrorLogging("ProcessSMTMO-Exception", LoginData.User, ex.Message & ex.Source, "E")
    '            ProcessSMTMO = ex.Message
    '        End Try

    '    End Using

    'End Function


    'Public Function TransactSMTMO(ByVal LoginData As ERPLogin, ByVal ds As DataSet) As String
    '    Dim da As DataAccess = GetDataAccess()
    '    Using connection As New OracleConnection(da._OConnString)
    '        Dim MO As String = ds.Tables(0).Rows(0)("MO").ToString

    '        Try
    '            connection.Open()
    '            Dim comm As OracleCommand = connection.CreateCommand()
    '            Dim OrgID As String = GetOrgID(LoginData.OrgCode)

    '            Try
    '                comm.CommandType = CommandType.StoredProcedure
    '                comm.CommandText = "apps.XXETR_wip_pkg.initialize"
    '                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
    '                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
    '                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
    '                comm.ExecuteOracleNonQuery("")
    '                comm.Parameters.Clear()

    '                Dim i As Integer
    '                Dim k As Integer = 1
    '                Dim DR() As DataRow

    '                'k=1:   First validate if there are any failed records
    '                'k=2:   Second commit if the all records are successful 
    '                While k < 3
    '                    For i = 0 To ds.Tables(0).Rows.Count - 1
    '                        comm.Parameters.Clear()
    '                        comm.Transaction = connection.BeginTransaction()
    '                        comm.CommandType = CommandType.StoredProcedure
    '                        'comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo"
    '                        comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo_by_lotnum"
    '                        comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
    '                        comm.Parameters.Add("p_org_id", OracleType.VarChar, 20).Value = OrgID                        'LoginData.OrgCode
    '                        comm.Parameters.Add("p_move_order", OracleType.VarChar, 50).Value = MO
    '                        comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
    '                        comm.Parameters.Add("p_lot_num", OracleType.VarChar, 50)
    '                        comm.Parameters.Add("p_quantity", OracleType.Double)
    '                        comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
    '                        comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)
    '                        comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
    '                        comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

    '                        comm.Parameters("p_item_num").Value = ds.Tables(0).Rows(i)("Item")
    '                        comm.Parameters("p_lot_num").Value = ds.Tables(0).Rows(i)("RTLot")
    '                        comm.Parameters("p_quantity").Value = ds.Tables(0).Rows(i)("Qty")
    '                        comm.Parameters("o_success_flag").Value = ds.Tables(0).Rows(i)("Status")
    '                        comm.Parameters("o_error_mssg").Value = ds.Tables(0).Rows(i)("Message")
    '                        comm.ExecuteNonQuery()

    '                        ds.Tables(0).Rows(i)("Status") = comm.Parameters("o_success_flag").Value
    '                        ds.Tables(0).Rows(i)("Message") = comm.Parameters("o_error_mssg").Value

    '                        If k = 1 Then
    '                            comm.Transaction.Rollback()
    '                        ElseIf k = 2 Then
    '                            comm.Transaction.Commit()
    '                        End If
    '                    Next

    '                    If k = 1 Then
    '                        DR = ds.Tables("MatlList").Select(" Status = 'N' or Status = ' ' or Status IS Null ")
    '                        If DR.Length > 0 Then k = 5 'No unsuccessful info is returned
    '                    End If

    '                    k = k + 1
    '                End While
    '                connection.Close()
    '                'If k = 3 Then Return "Y"

    '                DR = Nothing
    '                DR = ds.Tables("MatlList").Select(" Status = 'N' or Status = ' ' or Status IS Null ")
    '                If DR.Length = 0 Then
    '                    TransactSMTMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
    '                    Return TransactSMTMO
    '                End If

    '                For i = 0 To DR.Length - 1
    '                    Dim ErrMsg As String = ""
    '                    ErrMsg = "MO " & MO & " for Item: " & DR(i)("Item").ToString & " and RTLot " & DR(i)("RTLot").ToString & " with IssuedQty " & DR(i)("Qty").ToString & " and flag: " & DR(i)("Status").ToString & " and error message; " & DR(i)("Message").ToString
    '                    ErrorLogging("WMSMO-TransactSMTMO1", LoginData.User, ErrMsg, "I")
    '                Next

    '                TransactSMTMO = DR(0)("Message").ToString
    '                If TransactSMTMO = "" Then
    '                    TransactSMTMO = "No data is submitted in Oracle"
    '                    Return TransactSMTMO
    '                End If

    '                If TransactSMTMO.ToUpper.Contains("ORA-00054") Then
    '                    TransactSMTMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(0)("Item").ToString
    '                Else
    '                    TransactSMTMO = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
    '                End If
    '                Return TransactSMTMO                             'Return Error message
    '            Catch ex As Exception
    '                ErrorLogging("WMSMO-TransactSMTMO", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
    '                TransactSMTMO = ex.Message
    '            Finally
    '                If connection.State <> ConnectionState.Closed Then connection.Close()
    '            End Try

    '        Catch ex As Exception
    '            ErrorLogging("TransactSMTMO-Exception", LoginData.User, ex.Message & ex.Source, "E")
    '            TransactSMTMO = ex.Message
    '        End Try

    '    End Using

    'End Function

    Public Function ProcessSMTMO(ByVal LoginData As ERPLogin, ByVal ds As DataSet, ByRef mySMOData As SMTData) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim MO As String = ds.Tables(0).Rows(0)("MO").ToString
                Dim EventID As String = ds.Tables(0).Rows(0)("DJ").ToString               'Save EventID in DJ field here if ActionType = "MO Confirm"

                Dim oda As OracleDataAdapter
                Dim comm As OracleCommand

                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Try
                    oda = New OracleDataAdapter
                    comm = da.Ora_Command_Trans()
                Catch ex As Exception
                    ErrorLogging("ProcessSMTMO-OpenException", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                    ProcessSMTMO = ex.Message
                    Return ProcessSMTMO
                    Exit Function
                End Try

                Try
                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.XXETR_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                    comm.ExecuteOracleNonQuery("")
                    comm.Parameters.Clear()

                    comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo"
                    'comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo_by_lotnum"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                    comm.Parameters.Add("p_org_id", OracleType.VarChar, 20).Value = OrgID                        'LoginData.OrgCode
                    comm.Parameters.Add("p_move_order", OracleType.VarChar, 50).Value = MO
                    comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                    comm.Parameters.Add("p_quantity", OracleType.Double)
                    comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
                    comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

                    comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                    comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                    comm.Parameters("p_item_num").SourceColumn = "Item"
                    comm.Parameters("p_quantity").SourceColumn = "Qty"

                    comm.Parameters("o_success_flag").SourceColumn = "Status"
                    comm.Parameters("o_error_mssg").SourceColumn = "Message"

                    oda.InsertCommand = comm
                    oda.Update(ds.Tables(0))

                    If ds.Tables(0).Rows.Count = 0 Then
                        comm.Transaction.Rollback()
                        comm.Connection.Close()
                        ProcessSMTMO = "No data is submitted in Oracle"
                        Return ProcessSMTMO
                    End If

                    Dim DR() As DataRow = Nothing
                    DR = ds.Tables(0).Select(" Status = 'N' or Status = ' ' or Status IS Null ")
                    If DR.Length = 0 Then
                        comm.Transaction.Commit()
                        comm.Connection.Close()
                        ProcessSMTMO = "Y"                                'Oracle Data update successfully and set Flag = "Y"
                        Return ProcessSMTMO
                    End If

                    'Rollback and record error message if Oracle Data update failed
                    comm.Transaction.Rollback()
                    comm.Connection.Close()

                    'Read PickExactFlag in Event Header table for later use  -- 3/1/2018
                    Dim Sqlstr As String
                    Dim PickExactFlag As String = ""
                    If EventID <> "" Then
                        Sqlstr = String.Format("Select PickExactFlag from T_InvEventHeader with (nolock) where EventID = '{0}'", EventID)
                        PickExactFlag = Convert.ToString(da.ExecuteScalar(Sqlstr))
                    End If


                    Dim i As Integer
                    Dim ReAllocationMsg As String = ""
                    Dim ReAllocationFlag As Boolean = False
                    ProcessSMTMO = ""

                    For i = 0 To DR.Length - 1
                        Dim RtnMsg As String = DR(i)("Message").ToString

                        Dim ErrMsg As String = ""
                        ErrMsg = "MO " & MO & " for Item: " & DR(i)("Item").ToString & " with IssuedQty " & DR(i)("Qty").ToString & " and flag: " & DR(i)("Status").ToString & " and error message; " & RtnMsg
                        ErrorLogging("WMSMO-ProcessSMTMO1", LoginData.User, ErrMsg, "I")

                        'Check if there has user locked MO, if yes, give error message
                        If RtnMsg.Contains("ORA-00054") Then
                            'ProcessSMTMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(i)("Item").ToString
                            ProcessSMTMO = "^WMS-52@ " & MO & " ^WMS-53@ " & DR(i)("Item").ToString & "^WMS-54@"
                        ElseIf RtnMsg.Contains("not match oracle allocated qty") Then
                            'Check if there has error message which IssuedQty does not match with allocated qty, 
                            'and if PickExactFlag = “Y”, return error message only, otherwise call RedoLedMOAllocated function -- 3/1/2018
                            If PickExactFlag = "N" Then
                                ReAllocationFlag = True
                                ReAllocationMsg = ErrMsg
                            Else
                                If MissingPickExactCLID(mySMOData.EventID) = "RESUBMIT" Then
                                    ReAllocationFlag = True
                                    ReAllocationMsg = ErrMsg
                                Else
                                ProcessSMTMO = "MO " & MO & " ^WMS-55@ " & DR(i)("Item").ToString & " ^WMS-56@"
                                End If
                                Exit For
                            End If
                        ElseIf RtnMsg.Contains("tran_header_id is  ,") Then
                            'Check if there has error message which the column tran_header_id is blank, this case happened on 10/26/2016
                            ReAllocationFlag = True
                            ReAllocationMsg = ErrMsg
                        End If
                    Next

                    'Give error message if one user locked MO
                    If ProcessSMTMO <> "" Then Return ProcessSMTMO


                    'If ReAllocationFlag = True for MO Confirm only, then will call function to do reallocation first           -- 10/07/2016
                    If ReAllocationFlag = True AndAlso EventID <> "" Then
                        'Return Error Message if Reallocation still has problem         -- 11/24/2016
                        If mySMOData.CheckDock <> "" Then
                            ErrorLogging("WMSMO-ProcessSMTMO2", LoginData.User, "RedoLedMOAllocated error, " & ReAllocationMsg, "I")
                            Return ReAllocationMsg
                        End If


                        Try
                            ProcessSMTMO = RedoLedMOAllocated(EventID, LoginData)
                        Catch ex As Exception
                            ErrorLogging("ProcessSMTMO-call-RedoLedMOAllocated", LoginData.User, ex.Message & ex.Source, "E")
                            ProcessSMTMO = ex.Message
                        End Try
                        Return ProcessSMTMO
                    End If


                    ProcessSMTMO = DR(0)("Message").ToString
                    If ProcessSMTMO = "" Then
                        ProcessSMTMO = "No data is submitted in Oracle"
                    Else
                        ProcessSMTMO = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
                    End If
                    Return ProcessSMTMO


                    'If ProcessSMTMO.ToUpper.Contains("ORA-00054") Then
                    '    ProcessSMTMO = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(0)("Item").ToString
                    'Else
                    '    ProcessSMTMO = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
                    'End If
                    'Return ProcessSMTMO                                'Return Error message

                Catch ex As Exception
                    ErrorLogging("WMSMO-ProcessSMTMO", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                    comm.Transaction.Rollback()
                    ProcessSMTMO = ex.Message
                Finally
                    If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
                End Try

            Catch ex As Exception
                ErrorLogging("ProcessSMTMO-Exception", LoginData.User, ex.Message & ex.Source, "E")
                ProcessSMTMO = ex.Message
            End Try

        End Using

    End Function

    Public Function ProcessMissItem(ByVal LoginData As ERPLogin, ByRef ds As DataSet) As String
        Dim da As DataAccess = GetDataAccess()
        Using connection As New OracleConnection(da._OConnString)
            Dim MO As String = ds.Tables(0).Rows(0)("MO").ToString

            Try
                connection.Open()
                Dim comm As OracleCommand = connection.CreateCommand()
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Try
                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.XXETR_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID) '15904
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv   '54050
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                    comm.ExecuteOracleNonQuery("")
                    comm.Parameters.Clear()

                    Dim i As Integer
                    Dim k As Integer = 1
                    Dim DR() As DataRow

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        comm.Parameters.Clear()
                        comm.Transaction = connection.BeginTransaction()
                        comm.CommandType = CommandType.StoredProcedure
                        'comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo"
                        comm.CommandText = "apps.XXETR_wip_pkg.process_led_mo_by_lotnum"
                        comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                        comm.Parameters.Add("p_org_id", OracleType.VarChar, 20).Value = OrgID                        'LoginData.OrgCode
                        comm.Parameters.Add("p_move_order", OracleType.VarChar, 50).Value = MO
                        comm.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                        comm.Parameters.Add("p_lot_num", OracleType.VarChar, 50)
                        comm.Parameters.Add("p_quantity", OracleType.Double)
                        comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
                        comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)
                        comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                        comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                        comm.Parameters("p_item_num").Value = ds.Tables(0).Rows(i)("Item")
                        comm.Parameters("p_lot_num").Value = ds.Tables(0).Rows(i)("RTLot")
                        comm.Parameters("p_quantity").Value = ds.Tables(0).Rows(i)("Qty")
                        comm.Parameters("o_success_flag").Value = ds.Tables(0).Rows(i)("Status")
                        comm.Parameters("o_error_mssg").Value = ds.Tables(0).Rows(i)("Message")
                        comm.ExecuteNonQuery()

                        ds.Tables(0).Rows(i)("Status") = comm.Parameters("o_success_flag").Value
                        ds.Tables(0).Rows(i)("Message") = comm.Parameters("o_error_mssg").Value

                        If ds.Tables(0).Rows(i)("Status") = "N" Then
                            comm.Transaction.Rollback()
                            Exit For                                                              'Exit For if there has any error
                        Else
                            comm.Transaction.Commit()
                        End If
                    Next
                    connection.Close()

                    DR = Nothing
                    'DR = ds.Tables(0).Select(" Status = 'N' or Status = ' ' or Status IS Null ")
                    DR = ds.Tables(0).Select(" Status = 'N' ")
                    If DR.Length = 0 Then
                        ProcessMissItem = "Y"                                'Oracle Data update successfully and set Flag = "Y"
                        Return ProcessMissItem
                    End If

                    For i = 0 To DR.Length - 1
                        Dim ErrMsg As String = ""
                        ErrMsg = "MO " & MO & " for Item: " & DR(i)("Item").ToString & " and RTLot " & DR(i)("RTLot").ToString & " with IssuedQty " & DR(i)("Qty").ToString & " and flag: " & DR(i)("Status").ToString & " and error message; " & DR(i)("Message").ToString
                        ErrorLogging("WMSMO-ProcessMissItem1", LoginData.User, ErrMsg, "I")
                    Next

                    ProcessMissItem = DR(0)("Message").ToString
                    If ProcessMissItem = "" Then
                        ProcessMissItem = "No data is submitted in Oracle"
                        Return ProcessMissItem
                    End If

                    If ProcessMissItem.ToUpper.Contains("ORA-00054") Then
                        ProcessMissItem = "One user opened Transact Move Order in Oracle for MO: " & MO & " with Item: " & DR(0)("Item").ToString
                    Else
                        ProcessMissItem = "MO " & MO & " for Item: " & DR(0)("Item").ToString & " with error message; " & DR(0)("Message").ToString
                    End If
                    Return ProcessMissItem                             'Return Error message
                Catch ex As Exception
                    ErrorLogging("WMSMO-ProcessMissItem", LoginData.User, "MO " & MO & ", " & ex.Message & ex.Source, "E")
                    ProcessMissItem = ex.Message
                Finally
                    If connection.State <> ConnectionState.Closed Then connection.Close()
                End Try

            Catch ex As Exception
                ErrorLogging("ProcessMissItem-Exception", LoginData.User, ex.Message & ex.Source, "E")
                ProcessMissItem = ex.Message
            End Try

        End Using

    End Function

    Public Function RecordSMTCOO(ByVal LoginData As ERPLogin, ByVal dtCLIDs As DataTable) As Boolean
        Using da As DataAccess = GetDataAccess()
            RecordSMTCOO = False

            'add a table to record the COO Data for RM item
            Dim COOData As DataTable
            Dim RMItems As New DataSet
            COOData = New Data.DataTable("COOData")
            COOData.Columns.Add(New Data.DataColumn("DN", System.Type.GetType("System.String")))
            COOData.Columns.Add(New Data.DataColumn("MO", System.Type.GetType("System.String")))
            COOData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
            COOData.Columns.Add(New Data.DataColumn("RTLot", System.Type.GetType("System.String")))
            COOData.Columns.Add(New Data.DataColumn("COO", System.Type.GetType("System.String")))
            COOData.Columns.Add(New Data.DataColumn("PickedQty", System.Type.GetType("System.Decimal")))
            RMItems.Tables.Add(COOData)

            Try
                Dim i As Integer
                Dim DR() As DataRow
                Dim myDR As DataRow

                For i = 0 To dtCLIDs.Rows.Count - 1
                    Dim Matl As String = dtCLIDs.Rows(i)("Item").ToString
                    Dim RTLot As String = dtCLIDs.Rows(i)("RTLot").ToString
                    Dim COO As String = dtCLIDs.Rows(i)("COO").ToString

                    DR = COOData.Select("MaterialNo='" & Matl & "' and RTLot='" & RTLot & "' and COO='" & COO & "'")
                    If DR.Length = 0 Then
                        myDR = COOData.NewRow              'Record the COO
                        myDR("DN") = dtCLIDs.Rows(i)("DJ").ToString
                        myDR("MO") = dtCLIDs.Rows(i)("MO").ToString
                        myDR("MaterialNo") = Matl
                        myDR("RTLot") = RTLot
                        myDR("COO") = COO
                        myDR("PickedQty") = dtCLIDs.Rows(i)("Qty")
                        COOData.Rows.Add(myDR)
                    Else
                        DR(0)("PickedQty") = DR(0)("PickedQty") + dtCLIDs.Rows(i)("Qty")
                        DR(0).AcceptChanges()
                    End If
                Next

                RecordSMTCOO = InsertShipCOO(RMItems, LoginData)

            Catch ex As Exception
                ErrorLogging("WMSMO-RecordSMTCOO", LoginData.User, ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using

    End Function

    Public Function ReadEventID(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ReadEventID = False
            mySMOData.RtnMsg = ""

            Dim ds As New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_WMSReadEventID '{0}', '{1}' ", mySMOData.EventID, mySMOData.ActionType)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Header", "Item"}
                ds = da.ExecuteDataSet(sql, tables)

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    mySMOData.RtnMsg = "^WMS-39@ " & mySMOData.EventID
                    Return False
                End If

                If mySMOData.ActionType = "Read" OrElse mySMOData.ActionType = "Docking" Then
                    Dim EventType = ds.Tables(0).Rows(0)("EventType").ToString
                    If mySMOData.ActionType = "Read" Then
                        'Give error message if not MO Confirm/SO Shipment Type.
                        If Not (EventType = "MO Confirm" OrElse EventType = "SO Shipment") Then
                            mySMOData.RtnMsg = "^WMS-36@"
                            Return False
                        End If
                    End If

                    'Check Event Status, if not Closed, give error message
                    If ds.Tables(0).Rows(0)("Status").ToString <> "Closed" Then
                        mySMOData.RtnMsg = "^WMS-24@"
                        Return False
                    End If

                    'Check Event Remarks, if MO Cancelled, give error message
                    If ds.Tables(0).Rows(0)("Remarks").ToString = "MO Cancelled" Then
                        mySMOData.RtnMsg = "^WMS-44@"
                        Return False
                    End If

                    If mySMOData.ActionType = "Docking" Then
                        'Check EventType first, if not MO Confirm, give error message
                        If EventType <> "MO Confirm" Then
                            mySMOData.RtnMsg = "^WMS-34@"
                            Return False
                        End If

                        'Check if there has Primary CLIDs, if not, give error message
                        Dim KeyCLID As Integer = ds.Tables(0).Rows(0)("TotalStation")
                        If KeyCLID = 0 Then
                            mySMOData.RtnMsg = "^WMS-35@"
                            Return False
                        End If

                        Dim ProdFloor As String = ds.Tables(0).Rows(0)("ProdFloor").ToString
                        Dim ConfigFloor As String = ds.Tables(0).Rows(0)("ConfigFloor").ToString

                        'Give Information if the Target ProdFloor for this EventID is not 2S
                        If ProdFloor <> "" AndAlso ConfigFloor <> "" Then
                            If ConfigFloor.Contains(ProdFloor) = False Then
                                mySMOData.RtnMsg = "^WMS-33@ " & ConfigFloor & ", " & "^WMS-32@"
                                Return False
                            End If
                        End If

                        Dim DR() As DataRow
                        DR = ds.Tables(1).Select("IsPrimary = 'True' and DockSlot = '' ")
                        If DR.Length = KeyCLID Then                     'Not Assign Docking Slot at all
                            'For PC Docking, the Rack is not Blank, so we can Assign new Docking and Docking Slot for the Closed EventID 
                            If mySMOData.UserProdLine <> "" Then
                                ProcessDocking(mySMOData)
                                ds.Tables(0).Rows(0)("Docking") = mySMOData.UserProdLine
                                ds.Tables(0).AcceptChanges()

                                Dim dtItem As New DataTable
                                dtItem = mySMOData.dsItem.Tables(0).Copy

                                'Refresh the Event Item Data after Assigned Docking Slot
                                ds.Tables(1).Clear()
                                ds.Tables(1).Merge(dtItem)
                            End If

                            'For HH Docking, the Rack is Blank, then Clear up the old Docking for the Closed EventID
                            If mySMOData.UserProdLine = "" Then
                                ProcessDocking(mySMOData)
                                ds.Tables(0).Rows(0)("Docking") = ""
                                ds.Tables(0).AcceptChanges()
                            End If

                        End If

                    End If
                End If

                ReadEventID = True
                mySMOData.dsItem = New DataSet
                mySMOData.dsItem = ds

            Catch ex As Exception
                ErrorLogging("WMSMO-ReadEventID", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Reading EventID with error."
                Return False
            End Try
        End Using

    End Function

    Public Function PrintSMTDJLabel(ByRef mySMOData As SMTData, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            PrintSMTDJLabel = False

            Dim dtData As New DataTable
            Dim LabelData As PickingLabel = New PickingLabel

            Try
                dtData = mySMOData.dsItem.Tables(0)

                Dim TmpDate As Date
                TmpDate = dtData.Rows(0)("EventDate")

                LabelData.OrgCode = mySMOData.OrgCode
                LabelData.EventID = mySMOData.EventID
                LabelData.EventDate = TmpDate.ToString("MM/dd/yyyy")                     'EventID Closed Date   
                LabelData.MONo = dtData.Rows(0)("MO").ToString                                  'MO Number
                LabelData.DJ = dtData.Rows(0)("DJ").ToString
                LabelData.Job = dtData.Rows(0)("Job").ToString
                LabelData.PCBA = dtData.Rows(0)("PCBA").ToString
                LabelData.DJRev = dtData.Rows(0)("DJRev").ToString
                LabelData.DJQty = dtData.Rows(0)("Qty").ToString
                LabelData.ProdLine = dtData.Rows(0)("ProdLine").ToString
                LabelData.Floor = dtData.Rows(0)("ProdFloor").ToString
                LabelData.CLIDCount = dtData.Rows(0)("CLIDCount").ToString
                LabelData.PanelSide = dtData.Rows(0)("PanelSide").ToString
                LabelData.GroupFlag = dtData.Rows(0)("GroupFlag").ToString

                Dim sql As String
                Dim strContent As String
                Dim arryFile() As String

                arryFile = Split(PickingLabelFile, "\")
                strContent = "OrgCode^" & LabelData.OrgCode & "^EventID^" & LabelData.EventID & "^EventDate^" & LabelData.EventDate & "^MONo^" & LabelData.MONo & "^DJ^" & LabelData.DJ _
                                 & "^Job^" & LabelData.Job & "^PCBA^" & LabelData.PCBA & "^DJRev^" & LabelData.DJRev & "^DJQty^" & LabelData.DJQty & "^ProdLine^" & LabelData.ProdLine _
                                 & "^Floor^" & LabelData.Floor & "^CLIDCount^" & LabelData.CLIDCount & "^PanelSide^" & LabelData.PanelSide & "^GroupFlag^" & LabelData.GroupFlag

                sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                da.ExecuteScalar(sql)
                PrintSMTDJLabel = True

            Catch ex As Exception
                ErrorLogging("WMSMO-ReadEventID", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Print Picking Label with error."
                Return False
            End Try
        End Using

    End Function

    Public Function CancelMO(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            CancelMO = False
            mySMOData.RtnMsg = ""

            'For MO Cancel, only Change Event Header Status to Closed
            Dim Sqlstr, myResult As String
            Dim ActionType As String = "Cancel"
            Dim ErrMsg As String = "Cancel MO with error for MO " & mySMOData.MONo

            Try
                Dim CLID As String = ""                       'Leave CLID as Blank
                Dim dsSlot As New DataSet
                Sqlstr = String.Format("exec sp_WMSReadEventCLID '{0}', '{1}' ", mySMOData.EventID, CLID)
                dsSlot = da.ExecuteDataSet(Sqlstr, "dtSlot")

                If dsSlot Is Nothing OrElse dsSlot.Tables.Count = 0 OrElse dsSlot.Tables(0).Rows.Count = 0 Then
                ElseIf dsSlot.Tables(0).Rows.Count > 0 Then
                    '==Code--0/1/2(Off/On/blink. data type: integer)
                    '==Interval--time(data type: integer. 0:not limit)
                    Dim Code As Integer = 0
                    Dim SlotLightOn As Boolean = False

                    Dim myWMS As WMS = New WMS
                    dsSlot.DataSetName = "DS"
                    SlotLightOn = myWMS.LEDControlBySlot(dsSlot, Code, 3)
                End If


                'Cancel EventID, and change Event Header Status to Closed
                Sqlstr = String.Format("exec sp_WMSConfirmMO '{0}', '{1}', '{2}', N'{3}' ", mySMOData.OrgCode, mySMOData.EventID, ActionType, mySMOData.User)
                myResult = da.ExecuteScalar(Sqlstr).ToString
                If myResult = "Y" Then
                    CancelMO = True
                Else
                    mySMOData.RtnMsg = ErrMsg
                End If

            Catch ex As Exception
                ErrorLogging("WMSMO-CancelMO", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

    Public Function ProcessDocking(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ProcessDocking = False
            mySMOData.RtnMsg = ""

            Dim Sqlstr As String
            Dim ds As New DataSet
            Dim ErrMsg As String = "Process Docking with error for EventID " & mySMOData.EventID
            If mySMOData.EventID = "" Then ErrMsg = "Check Docking with error for Rack " & mySMOData.UserProdLine

            Try
                Sqlstr = String.Format("exec sp_WMSProcessDocking '{0}', '{1}', '{2}', N'{3}' ", mySMOData.EventID, mySMOData.OrgCode, mySMOData.UserProdLine, mySMOData.User)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Msg", "Item"}
                ds = da.ExecuteDataSet(sql, tables)

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    mySMOData.RtnMsg = "^WMS-39@ " & mySMOData.EventID
                    Return False
                End If

                Dim MsgFlag As String
                MsgFlag = ds.Tables(0).Rows(0)(0).ToString
                mySMOData.RtnMsg = ds.Tables(0).Rows(0)(1).ToString
                If MsgFlag = "N" Then Return False

                ProcessDocking = True
                mySMOData.dsItem = New DataSet
                mySMOData.dsItem.Merge(ds.Tables(1))

            Catch ex As Exception
                ErrorLogging("WMSMO-ProcessDocking", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

#End Region


#Region "LD-SMTMOConfirm"
    Public Function CheckProdLine_LD(ByRef mySMOData As SMTData) As String
        Using da As DataAccess = GetDataAccess()
            CheckProdLine_LD = ""

            Dim Sqlstr As String
            Dim ds As New DataSet

            Try

                Sqlstr = String.Format("exec sp_WMSCheckDock_LD '{0}', '{1}', N'{2}', '{3}' ", mySMOData.UserProdLine, mySMOData.OrgCode, mySMOData.User, mySMOData.CheckDock)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Msg", "Header", "Item", "dtBag"}
                ds = da.ExecuteDataSet(sql, tables)

                CheckProdLine_LD = ds.Tables(0).Rows(0)(0).ToString
                mySMOData.RtnMsg = ds.Tables(0).Rows(0)(1).ToString

                If CheckProdLine_LD = "Y" AndAlso ds.Tables(1).Rows.Count > 0 Then
                    Dim DJNo As String = ds.Tables(1).Rows(0)("DJ").ToString
                    Dim CHFlag As String = GetConfigValue("CLID014")

                    If CHFlag = "YES" AndAlso DJNo <> "" Then
                        'Read DJ Status from HuaWei Web Service 
                        Dim DJStatus As String = ""
                        DJStatus = CheckCHDJ(DJNo, "")                         ' DJStatus = "0"  means OK

                        ds.Tables(1).Rows(0)("CHFlag") = CHFlag
                        ds.Tables(1).Rows(0)("DJStatus") = DJStatus
                        ds.Tables(1).Rows(0).AcceptChanges()
                    End If
                End If

                mySMOData.dsItem = New DataSet
                mySMOData.dsItem.Merge(ds.Tables(1))
                mySMOData.dsItem.Merge(ds.Tables(2))

                If ds.Tables.Count > 3 Then
                    mySMOData.dsItem.Merge(ds.Tables(3))
                End If

            Catch ex As Exception
                ErrorLogging("WMSMO-CheckProdLine_LD", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Check ProdLine with error."
                Return "N"
            End Try
        End Using

    End Function

    Public Function ConfirmCLID_LD(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ConfirmCLID_LD = False

            Dim ds As New DataSet

            Try
                Dim Sqlstr, myResult As String

                'Check if this CLID exists in T_CLMaster or not for ActionType = Scanned / Missing, this is for SMT Confirm CLID only
                If mySMOData.ActionType <> "Docking" Then
                    Sqlstr = String.Format("exec sp_WMSReadEventCLID '{0}', '{1}' ", mySMOData.EventID, mySMOData.CLID)
                    ds = da.ExecuteDataSet(Sqlstr, "CLID")
                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                        mySMOData.RtnMsg = "Coundn't find CLID " & mySMOData.CLID & " in T_CLMaster"
                        Return False
                    End If

                    'For EventType = "SO Shipment", need to check if the Raw Item has COO or not
                    If mySMOData.EventType = "SO Shipment" Then
                        Dim COO As String = ds.Tables(0).Rows(0)("COO").ToString
                        Dim Matl As String = ds.Tables(0).Rows(0)("Item").ToString
                        Dim ItemType As String = ds.Tables(0).Rows(0)("ItemType").ToString

                        Dim ErrMsg As String
                        If ItemType = "RM" AndAlso COO = "" Then
                            ErrMsg = "^WMS-37@ " & Matl & " ^WMS-38@"
                            mySMOData.RtnMsg = ErrMsg
                            Return False
                        End If
                    End If

                    'If Event Header Status is 'LightOn', then Light Off the LED
                    Dim InvSlot As String = ds.Tables(0).Rows(0)("InvSlot").ToString
                    Dim EvtStatus As String = ds.Tables(0).Rows(0)("EvtStatus").ToString

                    Dim LightOn As Boolean = False
                    If EvtStatus = "LightOn" AndAlso InvSlot <> "" Then
                        'SlotLightOn(InvSlot, LightOn, mySMOData.User)

                        Dim dsSlot As New DataSet("DS")
                        Dim myDR As DataRow
                        Dim dtSlot As DataTable
                        dtSlot = New DataTable("dtSlot")
                        dtSlot.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                        dsSlot.Tables.Add(dtSlot)

                        myDR = dtSlot.NewRow()
                        myDR("slot") = InvSlot
                        dtSlot.Rows.Add(myDR)

                        '==Code--0/1/2(Off/On/blink. data type: integer)
                        '==Interval--time(data type: integer. 0:not limit)
                        Dim Code As Integer = 0
                        Dim myWMS As WMS = New WMS
                        myWMS.LEDControlBySlot(dsSlot, Code, 5)
                    End If
                End If


                'For SMT Confirm CLID with ActionType = Scanned / Missing, we need to update CLID Status here
                'For PC Docking to Scan Non-Primary CLID, we need to Update CLIDBag for Non-Primary CLID here
                Sqlstr = String.Format("exec sp_WMSConfirmCLID_LD '{0}', '{1}', '{2}', '{3}', N'{4}' ", mySMOData.OrgCode, mySMOData.EventID, mySMOData.CLID, mySMOData.ActionType, mySMOData.User)
                myResult = da.ExecuteScalar(Sqlstr).ToString

                If Left(myResult, 2) <> "Y/" Then
                    mySMOData.RtnMsg = myResult
                    Return False
                End If

                ConfirmCLID_LD = True
                Dim CLIDLists As String = Microsoft.VisualBasic.Mid(myResult, 3)
                mySMOData.RtnMsg = CLIDLists

                'No need to Light on the LED as there is no light here right now   -- 10/4/2017
                ''Light on the LED for the Slot if this CLID is Primary and it was not Splitted
                'Dim DockSlot As String
                'DockSlot = ds.Tables(0).Rows(0)("DockSlot").ToString
                'If DockSlot <> "" AndAlso CLIDLists = "" Then
                '    LightOn = True
                '    Dim Interval As Integer = 8                                              'User required to LightOn 8 seconds for Docking
                '    'SlotLightOn(DockSlot, LightOn, mySMOData.User)
                '    DockSlotLightOn(DockSlot, LightOn, mySMOData.User, Interval)
                'End If

            Catch ex As Exception
                ErrorLogging("WMSMO-ConfirmCLID_LD", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Confirm CLID with error."
                Return False
            End Try
        End Using

    End Function

    Public Function ReadEventID_LD(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ReadEventID_LD = False
            mySMOData.RtnMsg = ""

            Dim ds As New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_WMSReadEventID_LD '{0}', '{1}' ", mySMOData.EventID, mySMOData.ActionType)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Header", "Item", "dtBag"}
                ds = da.ExecuteDataSet(sql, tables)

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    mySMOData.RtnMsg = "^WMS-39@ " & mySMOData.EventID
                    Return False
                End If

                If mySMOData.ActionType = "Read" OrElse mySMOData.ActionType = "Docking" Then
                    Dim EventType = ds.Tables(0).Rows(0)("EventType").ToString
                    If mySMOData.ActionType = "Read" Then
                        'Give error message if not MO Confirm/SO Shipment Type.
                        If Not (EventType = "MO Confirm" OrElse EventType = "SO Shipment") Then
                            mySMOData.RtnMsg = "^WMS-36@"
                            Return False
                        End If
                    End If

                    'Check Event Status, if not Closed, give error message
                    If ds.Tables(0).Rows(0)("Status").ToString <> "Closed" Then
                        mySMOData.RtnMsg = "^WMS-24@"
                        Return False
                    End If

                    'Check Event Remarks, if MO Cancelled, give error message
                    If ds.Tables(0).Rows(0)("Remarks").ToString = "MO Cancelled" Then
                        mySMOData.RtnMsg = "^WMS-44@"
                        Return False
                    End If

                    If mySMOData.ActionType = "Docking" Then
                        'Check EventType first, if not MO Confirm, give error message
                        If EventType <> "MO Confirm" Then
                            mySMOData.RtnMsg = "^WMS-34@"
                            Return False
                        End If

                        'Check if there has Primary CLIDs, if not, give error message
                        Dim KeyCLID As Integer = ds.Tables(0).Rows(0)("TotalStation")
                        If KeyCLID = 0 Then
                            mySMOData.RtnMsg = "^WMS-35@"
                            Return False
                        End If

                        Dim ProdFloor As String = ds.Tables(0).Rows(0)("ProdFloor").ToString
                        Dim ConfigFloor As String = ds.Tables(0).Rows(0)("ConfigFloor").ToString

                        'Give Information if the Target ProdFloor for this EventID is not 2S
                        If ProdFloor <> "" AndAlso ConfigFloor <> "" Then
                            If ConfigFloor.Contains(ProdFloor) = False Then
                                mySMOData.RtnMsg = "^WMS-33@ " & ConfigFloor & ", " & "^WMS-32@"
                                Return False
                            End If
                        End If

                        Dim DR() As DataRow
                        DR = ds.Tables(1).Select("IsPrimary = 'True' and DockSlot = '' ")
                        If DR.Length = KeyCLID Then                     'Not Assign Docking Slot at all
                            'For PC Docking, the Rack is not Blank, so we can Assign new Docking and Docking Slot for the Closed EventID 
                            If mySMOData.UserProdLine <> "" Then
                                ProcessDocking_LD(mySMOData)
                                ds.Tables(0).Rows(0)("Docking") = mySMOData.UserProdLine
                                ds.Tables(0).AcceptChanges()

                                Dim dtItem As New DataTable
                                dtItem = mySMOData.dsItem.Tables(0).Copy

                                'Refresh the Event Item Data after Assigned Docking Slot
                                ds.Tables(1).Clear()
                                ds.Tables(1).Merge(dtItem)

                                Dim dtBag As New DataTable
                                dtBag = mySMOData.dsItem.Tables(1).Copy

                                'The first time to Assign Bag Header, then need to print Spare Label
                                If dtBag.Rows.Count > 0 Then
                                    ds.Tables(0).Rows(0)("AutoSparePT") = "Y"
                                    ds.Tables(0).AcceptChanges()
                                End If

                                'Refresh the Bag Header Data after Assigned Bags
                                ds.Tables(2).Clear()
                                ds.Tables(2).Merge(dtBag)
                            End If

                            'For HH Docking, the Rack is Blank, then Clear up the old Docking for the Closed EventID
                            If mySMOData.UserProdLine = "" Then
                                ProcessDocking_LD(mySMOData)
                                ds.Tables(0).Rows(0)("Docking") = ""
                                ds.Tables(0).AcceptChanges()
                            End If

                        End If

                    End If
                End If

                ReadEventID_LD = True
                mySMOData.dsItem = New DataSet
                mySMOData.dsItem = ds

            Catch ex As Exception
                ErrorLogging("WMSMO-ReadEventID_LD", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = "Reading EventID with error."
                Return False
            End Try
        End Using

    End Function

    Public Function PrintSMTDJLabel_LD(ByRef mySMOData As SMTData, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            PrintSMTDJLabel_LD = False

            Dim dtData As New DataTable
            Dim LabelData As LDPickingLabel = New LDPickingLabel

            Dim LabelType As String = mySMOData.ActionType                     ' P: Picking Label  / S: Spare Label
            Dim LabelSpare As LDSpareLabel = New LDSpareLabel

            Dim sql As String
            Dim strContent As String
            Dim arryFile() As String

            Dim ErrMsg As String = "Print Picking Label with error."

            Try
                dtData = mySMOData.dsItem.Tables(0)

                'P: Picking Label  -- Print Picking Label here
                If LabelType = "P" Then
                    Dim TmpDate As Date
                    TmpDate = dtData.Rows(0)("EventDate")

                    LabelData.OrgCode = mySMOData.OrgCode
                    LabelData.EventID = mySMOData.EventID
                    LabelData.EventDate = TmpDate.ToString("MM/dd/yyyy")                     'EventID Closed Date   
                    LabelData.MONo = dtData.Rows(0)("MO").ToString                           'MO Number
                    LabelData.DJ = dtData.Rows(0)("DJ").ToString
                    LabelData.Job = dtData.Rows(0)("Job").ToString
                    LabelData.PCBA = dtData.Rows(0)("PCBA").ToString
                    LabelData.DJRev = dtData.Rows(0)("DJRev").ToString
                    LabelData.DJQty = dtData.Rows(0)("Qty").ToString
                    LabelData.ProdLine1 = dtData.Rows(0)("Line1").ToString
                    LabelData.ProdLine2 = dtData.Rows(0)("Line2").ToString
                    LabelData.ProdLine3 = dtData.Rows(0)("Line3").ToString
                    LabelData.Floor = dtData.Rows(0)("ProdFloor").ToString
                    LabelData.CLIDCount = dtData.Rows(0)("CLIDCount").ToString

                    arryFile = Split(LDPickingLabelFile, "\")
                    strContent = "OrgCode^" & LabelData.OrgCode & "^EventID^" & LabelData.EventID & "^EventDate^" & LabelData.EventDate & "^MONo^" & LabelData.MONo & "^DJ^" & LabelData.DJ _
                               & "^Job^" & LabelData.Job & "^PCBA^" & LabelData.PCBA & "^DJRev^" & LabelData.DJRev & "^DJQty^" & LabelData.DJQty & "^ProdLine1^" & LabelData.ProdLine1 _
                               & "^ProdLine2^" & LabelData.ProdLine2 & "^ProdLine3^" & LabelData.ProdLine3 & "^Floor^" & LabelData.Floor & "^CLIDCount^" & LabelData.CLIDCount

                    sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                    da.ExecuteScalar(sql)
                    Return True
                End If


                'S: Spare Label  -- Print Spare Label here
                Dim dtBag As New DataTable
                dtBag = mySMOData.dsItem.Tables(1)
                ErrMsg = "Print Spare Label with error."

                LabelSpare.OrgCode = mySMOData.OrgCode
                LabelSpare.EventID = mySMOData.EventID
                LabelSpare.MONo = dtData.Rows(0)("MO").ToString                                  'MO Number
                LabelSpare.DJ = dtData.Rows(0)("DJ").ToString
                LabelSpare.Job = dtData.Rows(0)("Job").ToString
                LabelSpare.PCBA = dtData.Rows(0)("PCBA").ToString
                LabelSpare.DJQty = dtData.Rows(0)("Qty").ToString


                Dim i, j As Integer
                For i = 0 To dtBag.Rows.Count - 1
                    Dim myLine As String = ""
                    Dim myPanel As String = ""
                    Dim ProdLine As String = dtBag.Rows(i)("ProdLine").ToString

                    If ProdLine <> "" Then
                        myLine = Microsoft.VisualBasic.Left(ProdLine, ProdLine.Length - 1)
                        myPanel = Microsoft.VisualBasic.Right(ProdLine, 1)
                    End If

                    LabelSpare.ProdLine = myLine
                    LabelSpare.Panel = myPanel

                    Dim k As Integer = CInt(dtBag.Rows(i)("NumOfBags"))
                    Dim GroupID As String = dtBag.Rows(i)("GroupID").ToString
                    LabelSpare.GroupBagID = GroupID

                    If k = 0 Then
                        arryFile = Split(LDSpareLabelFile, "\")
                        strContent = "OrgCode^" & LabelSpare.OrgCode & "^EventID^" & LabelSpare.EventID & "^MONo^" & LabelSpare.MONo & "^DJ^" & LabelSpare.DJ _
                                   & "^Job^" & LabelSpare.Job & "^PCBA^" & LabelSpare.PCBA & "^DJQty^" & LabelSpare.DJQty & "^GroupBagID^" & LabelSpare.GroupBagID _
                                   & "^ProdLine^" & LabelSpare.ProdLine & "^Panel^" & LabelSpare.Panel

                        sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                        da.ExecuteScalar(sql)
                        Sleep(5)
                    Else
                        For j = 1 To k
                            LabelSpare.GroupBagID = GroupID & "-" & j.ToString

                            arryFile = Split(LDSpareLabelFile, "\")
                            strContent = "OrgCode^" & LabelSpare.OrgCode & "^EventID^" & LabelSpare.EventID & "^MONo^" & LabelSpare.MONo & "^DJ^" & LabelSpare.DJ _
                                       & "^Job^" & LabelSpare.Job & "^PCBA^" & LabelSpare.PCBA & "^DJQty^" & LabelSpare.DJQty & "^GroupBagID^" & LabelSpare.GroupBagID _
                                       & "^ProdLine^" & LabelSpare.ProdLine & "^Panel^" & LabelSpare.Panel

                            sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                            da.ExecuteScalar(sql)
                            Sleep(5)
                        Next
                    End If
                Next
                PrintSMTDJLabel_LD = True

            Catch ex As Exception
                ErrorLogging("WMSMO-PrintSMTDJLabel_LD", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

    Public Function ProcessDocking_LD(ByRef mySMOData As SMTData) As Boolean
        Using da As DataAccess = GetDataAccess()
            ProcessDocking_LD = False
            mySMOData.RtnMsg = ""

            Dim Sqlstr As String
            Dim ds As New DataSet
            Dim ErrMsg As String = "Process Docking with error for EventID " & mySMOData.EventID
            If mySMOData.EventID = "" Then ErrMsg = "Check Docking with error for Rack " & mySMOData.UserProdLine

            Try
                Sqlstr = String.Format("exec sp_WMSProcessDocking_LD '{0}', '{1}', '{2}', N'{3}' ", mySMOData.EventID, mySMOData.OrgCode, mySMOData.UserProdLine, mySMOData.User)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"Msg", "Item", "dtBag"}
                ds = da.ExecuteDataSet(sql, tables)

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    mySMOData.RtnMsg = "^WMS-39@ " & mySMOData.EventID
                    Return False
                End If

                Dim MsgFlag As String
                MsgFlag = ds.Tables(0).Rows(0)(0).ToString
                mySMOData.RtnMsg = ds.Tables(0).Rows(0)(1).ToString
                If MsgFlag = "N" Then Return False

                ProcessDocking_LD = True
                mySMOData.dsItem = New DataSet
                mySMOData.dsItem.Merge(ds.Tables(1))
                mySMOData.dsItem.Merge(ds.Tables(2))

            Catch ex As Exception
                ErrorLogging("WMSMO-ProcessDocking_LD", mySMOData.User, ex.ToString, "E")
                mySMOData.RtnMsg = ErrMsg
                Return False
            End Try
        End Using

    End Function

#End Region

End Class
