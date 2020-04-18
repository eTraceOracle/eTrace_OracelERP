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


Public Class Putaway
    Inherits PublicFunction


#Region "--------Component to DJ Oracle_JIM2009"

    Public Function SourceForCompToDJ(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet     ' ByVal StorageType As String
        Try
            Using da As DataAccess = GetDataAccess()

                Dim Sqlstr As String
                Dim newColumn As DataColumn
                If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                    'Sqlstr = String.Format("Select BoxID as CLID, MaterialNo, MaterialRevision, SLOC, StorageBin, RTLot, SUM(QtyBaseUOM) AS QtyBaseUOM, BaseUOM from T_CLMaster where SLOC IS NOT NULL and StatusCode <> '0' and  BoxID = '{0}' and OrgCode = '{1}' group by BoxID, MaterialNo, MaterialRevision, SLOC, StorageBin, RTLot, BaseUOM", CLID, OracleLoginData.OrgCode)
                    Sqlstr = String.Format("Select CLID, MaterialNo, MaterialRevision, SLOC, StorageBin, RTLot, QtyBaseUOM, BaseUOM, BoxID from T_CLMaster with (nolock) where SLOC IS NOT NULL and StatusCode = '1' and  BoxID = '{0}' and OrgCode = '{1}'", CLID, OracleLoginData.OrgCode)
                    SourceForCompToDJ = da.ExecuteDataSet(Sqlstr, "SourceData")
                Else
                    Sqlstr = String.Format("Select CLID, MaterialNo, MaterialRevision, SLOC, StorageBin, RTLot, QtyBaseUOM, BaseUOM, BoxID from T_CLMaster with (nolock) where SLOC IS NOT NULL and StatusCode = '1' and  CLID = '{0}' and OrgCode = '{1}'", CLID, OracleLoginData.OrgCode)
                    SourceForCompToDJ = da.ExecuteDataSet(Sqlstr, "SourceData")
                End If
                '' Sqlstr = String.Format("Select Item,Material,MaterialRevision,ReqQty,PickedQty,Status from T_PDTOItem where PDTO = {0}", CLID)
                'Sqlstr = String.Format("Select CLID, MaterialNo, MaterialRevision, SLOC, StorageBin, RTLot, QtyBaseUOM, BaseUOM, MaterialRevision from T_CLMaster where SLOC IS NOT NULL and StatusCode <> '0' and  CLID = '{0}' and OrgCode = '{1}'", CLID, OracleLoginData.OrgCode)
                'SourceForCompToDJ = da.ExecuteDataSet(Sqlstr, "SourceData")

                newColumn = New DataColumn("DJNo", System.Type.GetType("System.String"))
                SourceForCompToDJ.Tables("SourceData").Columns.Add(newColumn)
                newColumn = New DataColumn("Reason", System.Type.GetType("System.String"))
                SourceForCompToDJ.Tables("SourceData").Columns.Add(newColumn)

                Return SourceForCompToDJ
            End Using
        Catch ex As Exception
            Throw ex
            ErrorLogging("Issue Component to DJ - SourceForCompToDJ", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function WS_CompToDJ(ByVal MatlList As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet     ' 
        Try
            WS_CompToDJ = New DataSet
            WS_CompToDJ = MatlList.Clone  'Cone table and structure, no data is copied

            Dim newColumn As DataColumn
            newColumn = New DataColumn("Org_code", System.Type.GetType("System.String"))
            WS_CompToDJ.Tables("SourceData").Columns.Add(newColumn)

            newColumn = New DataColumn("o_success_flag", System.Type.GetType("System.String"))
            WS_CompToDJ.Tables("SourceData").Columns.Add(newColumn)

            newColumn = New DataColumn("o_error_msgg", System.Type.GetType("System.String"))
            WS_CompToDJ.Tables("SourceData").Columns.Add(newColumn)

            '     ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", OracleLoginData.User, MatlList.Tables("SourceData").Rows.Count, "I")
            'ErrorLogging("PutAwayBinToBinPost", "JIM", "start to select from delivery", "I")
            For i As Integer = 0 To MatlList.Tables("SourceData").Rows.Count - 1
                'ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", OracleLoginData.User, i.ToString & "-MaterialNo = '" & MatlList.Tables("SourceData").Rows(i)("MaterialNo").ToString & _
                '                           "' and MaterialRevision = '" & MatlList.Tables("SourceData").Rows(i)("MaterialRevision").ToString & "' and SLOC = '" & MatlList.Tables("SourceData").Rows(i)("SLOC").ToString & _
                '                           "' and StorageBin = '" & MatlList.Tables("SourceData").Rows(i)("StorageBin").ToString & _
                '                           "' and RTLot = '" & MatlList.Tables("SourceData").Rows(i)("RTLot").ToString & "'", "I")
                Dim DR_of_matlList() As DataRow
                DR_of_matlList = Nothing
                DR_of_matlList = WS_CompToDJ.Tables("SourceData").Select("MaterialNo = '" & MatlList.Tables("SourceData").Rows(i)("MaterialNo").ToString & _
                                           "' and MaterialRevision = '" & MatlList.Tables("SourceData").Rows(i)("MaterialRevision").ToString & _
                                           "' and SLOC = '" & MatlList.Tables("SourceData").Rows(i)("SLOC").ToString & _
                                           "' and StorageBin = '" & MatlList.Tables("SourceData").Rows(i)("StorageBin").ToString & _
                                           "' and RTLot = '" & MatlList.Tables("SourceData").Rows(i)("RTLot").ToString & "'")
                'ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", "Jim", "DR_select", "I")
                '     ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", OracleLoginData.User, MatlList.Tables("SourceData").Rows(i)("MaterialNo") & MatlList.Tables("SourceData").Rows(i)("MaterialRevision"), "I")

                If DR_of_matlList.Length > 0 Then
                    DR_of_matlList(0)("QtyBaseUOM") = DR_of_matlList(0)("QtyBaseUOM") + MatlList.Tables("SourceData").Rows(i)("QtyBaseUOM")
                    'DR_of_matlList(0).AcceptChanges()
                    'ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", oraclelogindata.user, i.ToString & "Qty++", "I")
                Else
                    Dim dr As DataRow = WS_CompToDJ.Tables("SourceData").NewRow()
                    dr("MaterialNo") = MatlList.Tables("SourceData").Rows(i)("MaterialNo")
                    '     ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", OracleLoginData.User, dr("MaterialNo").ToString, "I")
                    dr("MaterialRevision") = MatlList.Tables("SourceData").Rows(i)("MaterialRevision")
                    '     ErrorLogging("eTracePutawaySourceForMatlMove(WS_CompToDJ)", OracleLoginData.User, dr("MaterialRevision").ToString, "I")
                    dr("SLOC") = MatlList.Tables("SourceData").Rows(i)("SLOC")
                    dr("StorageBin") = MatlList.Tables("SourceData").Rows(i)("StorageBin")
                    dr("RTLot") = MatlList.Tables("SourceData").Rows(i)("RTLot")
                    dr("QtyBaseUOM") = MatlList.Tables("SourceData").Rows(i)("QtyBaseUOM")
                    dr("BaseUOM") = MatlList.Tables("SourceData").Rows(i)("BaseUOM")
                    dr("DJNo") = MatlList.Tables("SourceData").Rows(0)("DJNo")
                    dr("MaterialRevision") = MatlList.Tables("SourceData").Rows(i)("MaterialRevision")
                    dr("Reason") = MatlList.Tables("SourceData").Rows(i)("Reason")
                    dr("Org_code") = OracleLoginData.OrgCode
                    WS_CompToDJ.Tables("SourceData").Rows.Add(dr)
                End If
            Next

            ''It is for debug
            'Dim iiii As Integer = 0
            'For Each dr As DataRow In WS_CompToDJ.Tables("SourceData").Rows
            '    iiii = iiii + 1
            '    Dim str_temp As String = iiii.ToString & "->" & dr("MaterialNo").ToString & " |" & _
            '     dr("SLOC").ToString & " |" & _
            '     dr("StorageBin").ToString & " |" & _
            '     dr("RTLot").ToString & " |" & _
            '     dr("QtyBaseUOM").ToString & " |" & _
            '     dr("BaseUOM").ToString & " |" & _
            '     dr("DJNo").ToString & " |" & _
            '     dr("MaterialRevision").ToString & " |" & _
            '     dr("Reason").ToString & " |" & _
            '     dr("Org_code").ToString
            '    'ErrorLogging("Putaway-SourceForMatlMove(WS_CompToDJ)", OracleLoginData.User.ToUpper, str_temp, "I")
            'Next
        Catch ex As Exception
            ErrorLogging("Issue Component to DJ - Collect Source for Post)", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
        End Try

        '*********************************************************************************************************************

        'If WS_CompToDJ.Tables("SourceData").Rows(0).RowState = DataRowState.Unchanged Then
        '    WS_CompToDJ.Tables("SourceData").Rows(0).SetAdded()
        'End If

        'ErrorLogging("Putaway-WSCOMP2DJ", OracleLoginData.User.ToUpper, WS_CompToDJ.Tables("SourceData").Rows.Count.ToString, "I")
        'ErrorLogging("Putaway-WSCOMP2DJ SourceData row status", OracleLoginData.User, WS_CompToDJ.Tables("SourceData").Rows(0).RowState.ToString(), "I")


        'ErrorLogging(WS + "WS_CompToDJ", OracleLoginData.User, "Dim oracle adapter is starting...", "I")

        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString

        Try

            Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

            'Dim resp_id As Integer = 54050
            'Dim appl_id As Integer = 706


            'ErrorLogging(WS + "WS_CompToDJ", OracleLoginData.User, "initialize oracle is starting...", "I")

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.initialize"
            comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)
            comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = OracleLoginData.RespID_Inv 'resp_id
            comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = OracleLoginData.AppID_Inv 'appl_id
            comm.ExecuteOracleNonQuery(aa)
            comm.Parameters.Clear()

            'ErrorLogging(WS + "WS_CompToDJ", OracleLoginData.User, "process_wip_issue oracle is starting...", "I")

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.process_wip_issue"

            comm.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 240))
            comm.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OrgID
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

            comm.Parameters("p_dj_name").SourceColumn = "DJNo"
            'comm.Parameters("p_org_code").SourceColumn = "Org_code"
            comm.Parameters("p_item_num").SourceColumn = "MaterialNo"
            comm.Parameters("p_item_rev").SourceColumn = "MaterialRevision"
            comm.Parameters("p_subinventory").SourceColumn = "SLOC"
            comm.Parameters("p_locator").SourceColumn = "StorageBin"
            comm.Parameters("p_lot_number").SourceColumn = "RTLot"
            comm.Parameters("p_issue_quantity").SourceColumn = "QtyBaseUOM"
            comm.Parameters("p_uom").SourceColumn = "BaseUOM"
            comm.Parameters("p_reason").SourceColumn = "Reason"
            comm.Parameters("o_success_flag").SourceColumn = "o_success_flag"
            comm.Parameters("o_error_msgg").SourceColumn = "o_error_msgg"

            oda_h.InsertCommand = comm
            oda_h.Update(WS_CompToDJ.Tables("SourceData"))

            'ErrorLogging(WS + "WS_CompToDJ", OracleLoginData.User, "process_wip_issue oracle is end", "I")

            Dim DR() As DataRow = Nothing
            DR = WS_CompToDJ.Tables("SourceData").Select("o_success_flag = 'N'")

            If DR.Length = 0 Then   'No unsuccessful info is returned

                comm.Transaction.Commit()
                comm.Connection.Close()
                comm.Connection.Dispose()
                comm.Dispose()
                'ErrorLogging("eTraceOraclePutaway", OracleLoginData.User, "Commit is succesful", "I")
            Else
                For jj As Integer = 0 To (DR.Length - 1)
                    ErrorLogging("Issue Component to DJ - Post", OracleLoginData.User, DR(jj)("MaterialNo").ToString & DR(jj)("MaterialRevision").ToString & DR(jj)("o_error_msgg").ToString, "I")
                Next

                comm.Transaction.Rollback()
                comm.Connection.Close()
                comm.Connection.Dispose()
                comm.Dispose()
                ErrorLogging("Issue Component to DJ - Post", OracleLoginData.User, "Rollback is triggered", "I")
                Return WS_CompToDJ  ' Return if there is error information returned from oracle


            End If


            '********************* Proceed updating eTrace database if no any error is returned from oracle 
            'Following code is going to update label ID

            Dim DJInform As New DataSet
            Using das As DataAccess = GetDataAccess()
                Dim oda As OracleDataAdapter = das.Oda_Sele()
                Try
                    Dim ds As New DataSet()
                    ds.Tables.Add("DJInfo")

                    oda.SelectCommand.CommandType = CommandType.StoredProcedure
                    oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_dj"

                    oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                    oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = MatlList.Tables("SourceData").Rows(0)("DJNo").ToString.Trim.ToUpper
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OracleLoginData.OrgCode

                    oda.SelectCommand.Connection.Open()
                    oda.Fill(ds, "DJInfo")
                    oda.SelectCommand.Connection.Close()
                    DJInform = ds
                Catch oe As Exception
                    Throw oe
                    ErrorLogging("Material replenishment-GetOrderInfoFromERP", "", "DJ: " & MatlList.Tables("SourceData").Rows(0)("DJNo").ToString.Trim.ToUpper & ", " & oe.Message & oe.Source, "E")
                Finally
                    If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
                End Try
            End Using

            Using da_update_eTrace As DataAccess = GetDataAccess()
                Dim sqlstr_T_PO_CLID As String
                Dim sqlstr_T_CLMaster As String
                Dim i As Integer
                Dim CLID As String
                Dim Qty As Double
                Dim DJ As String
                Dim Product As String
                Dim sqlstr, sqlCH1, InsertFlag, hwDJ As String

                'ErrorLogging("eTraceOraclePutaway-WS_CompToDJ", OracleLoginData.User.ToUpper, "Updating eTrace begin", "I")
                Try
                    DJ = MatlList.Tables("SourceData").Rows(0)("DJNo").ToString.Trim.ToUpper
                    Product = FixNull(DJInform.Tables(0).Rows(0)("product_number"))

                    For i = 0 To MatlList.Tables("SourceData").Rows.Count - 1
                        CLID = MatlList.Tables("SourceData").Rows(i)("CLID").ToString
                        Qty = Convert.ToDecimal(MatlList.Tables("SourceData").Rows(i)("QtyBaseUOM").ToString)
                        sqlstr_T_PO_CLID = String.Format("INSERT INTO T_PO_CLID(PO,Product,CLID,CLIDQty,ChangedBy,IssueDate,OrgCode) values ('{0}','{1}','{2}','{3}','{4}',getDate(),'{5}')", DJ, Product, CLID, Qty, OracleLoginData.User, OracleLoginData.OrgCode)
                        sqlstr_T_CLMaster = String.Format("UPDATE T_CLMaster set SupplyType='PUSH',statusCode='0',ChangedOn=getDate(),ChangedBy='{0}',LastDJ='{1}',StorageType='',LastTransaction='Component issue to DJ' where CLID='{2}'", OracleLoginData.User, DJ, CLID)
                        da_update_eTrace.ExecuteNonQuery(sqlstr_T_PO_CLID)
                        da_update_eTrace.ExecuteNonQuery(sqlstr_T_CLMaster)

                        sqlstr = String.Format("select top 1 Value from T_Config with (nolock) where ConfigID = 'CLID014'")
                        InsertFlag = da.ExecuteScalar(sqlstr)
                        If InsertFlag = "YES" Then
                            Dim DJ_Flag As String
                            Dim LDHWWS As CHWS.CollectDataToQDB = New CHWS.CollectDataToQDB
                            DJ_Flag = ""
                            DJ_Flag = LDHWWS.DJStatus(DJ)

                            If DJ_Flag = "0" Then
                                sqlCH1 = String.Format("INSERT INTO T_PO_CLID_CH1(PO,Product,CLID,CLIDQty,ChangedBy,IssueDate,OrgCode,Flag) values ('{0}','{1}','{2}','{3}','{4}',getDate(),'{5}',0)", DJ, Product, CLID, Qty, OracleLoginData.User, OracleLoginData.OrgCode)
                                da_update_eTrace.ExecuteNonQuery(sqlCH1)
                            End If

                            'sqlstr = String.Format("select top 1 ISNULL(Owner,'') as Owner FROM ldetraceaddition.DBHuaWei.dbo.T_Relation with (nolock) where ItemType = 'HWDJ' and Owner = '{0}'", DJ)
                            'hwDJ = da.ExecuteScalar(sqlstr)

                            'If FixNull(hwDJ) <> "" Then
                            '    sqlCH1 = String.Format("INSERT INTO T_PO_CLID_CH1(PO,Product,CLID,CLIDQty,ChangedBy,IssueDate,OrgCode,Flag) values ('{0}','{1}','{2}','{3}','{4}',getDate(),'{5}',0)", DJ, Product, CLID, Qty, OracleLoginData.User, OracleLoginData.OrgCode)
                            '    da_update_eTrace.ExecuteNonQuery(sqlCH1)
                            'End If
                        End If
                    Next

                    'ErrorLogging("eTraceOraclePutaway-WS_CompToDJ", OracleLoginData.User.ToUpper, "Updating eTrace end", "I")
                Catch ex As Exception
                    ErrorLogging("Issue Component to DJ - Update eTrace", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                End Try
            End Using

            '

        Catch ex As Exception
            'Dim aa As String = ex.Message
            oda_h.InsertCommand.Transaction.Rollback()
            ErrorLogging("Issue Component to DJ - Post", OracleLoginData.User.ToUpper, "exception: " & ex.Message & ex.Source, "E")
            Return WS_CompToDJ
        End Try
    End Function

#End Region


#Region "Putaway"

    Public Function GetValidSource(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal SubInv As String, ByVal Locator As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Try

                Dim ds = New DataSet
                Dim MaterialNo, RecDocNo, RTLot, PredefinedSubInv, PredefinedLocator As String
                Dim OrgCode As String = LoginData.OrgCode

                Dim Sqlstr As String
                Dim BoxFlag As Boolean = False
                If Mid(CLID, 3, 1) = "B" OrElse Left(CLID, 1) = "P" Then       'BoxID / PalletID
                    BoxFlag = True
                    Sqlstr = String.Format("Select TOP (1) CLID,MaterialNo,RecDocNo,RTLot,RecDate,StockType,PredefinedSubInv,PredefinedLocator,RecDocItem as TransactionID,BoxID, StatusCode, Status = '', Message = '' from T_CLMaster with (nolock) where (StatusCode = 1 or StatusCode = 9 ) and (SLOC IS NULL or SLOC = '') and BoxID = '{0}' and OrgCode = '{1}' ORDER BY StatusCode DESC, CLID ", CLID, LoginData.OrgCode)
                Else
                    Sqlstr = String.Format("Select CLID,MaterialNo,RecDocNo,RTLot,RecDate,StockType,PredefinedSubInv,PredefinedLocator,RecDocItem as TransactionID,BoxID, StatusCode, Status = '', Message = '' from T_CLMaster with (nolock) where (StatusCode = 1 or StatusCode = 9 ) and (SLOC IS NULL or SLOC = '') and CLID = '{0}' and OrgCode = '{1}'", CLID, LoginData.OrgCode)
                End If
                ds = da.ExecuteDataSet(Sqlstr, "CLIDS")

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    GetValidSource = Nothing
                    Exit Function
                End If

                'Check if there has CLID which has NOT done AML check yet
                If ds.Tables(0).Rows(0)("StatusCode") = "9" Then
                    Dim tmpCLID As String = ds.Tables(0).Rows(0)("CLID").ToString
                    ds.Tables(0).Rows(0)("Status") = "E"
                    If BoxFlag = True Then
                        ds.Tables(0).Rows(0)("Message") = "The BoxID contained one CLID " & tmpCLID & " which has not done AML check yet, please request receiving to do it."
                    Else
                        ds.Tables(0).Rows(0)("Message") = "This CLID " & tmpCLID & " has not done AML check yet, please request receiving to do it."
                    End If
                    Return ds
                End If

                MaterialNo = ds.Tables(0).Rows(0)("MaterialNo").ToString
                RecDocNo = ds.Tables(0).Rows(0)("RecDocNo").ToString
                RTLot = ds.Tables(0).Rows(0)("RTLot").ToString
                PredefinedSubInv = ds.Tables(0).Rows(0)("PredefinedSubInv").ToString
                PredefinedLocator = ds.Tables(0).Rows(0)("PredefinedLocator").ToString

                'If the BoxID has more than 500 CLIDs, then we only return BoxID to speed up the performance
                Dim CLIDCount As Integer = 0
                Dim ReturnBoxID As Boolean = False

                If BoxFlag = True Then
                    Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID002'")
                    CLIDCount = CInt(da.ExecuteScalar(Sqlstr))

                    ds = New DataSet
                    Sqlstr = String.Format("Select CLID from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and StatusCode = 1 and RecDocNo ='{0}'and MaterialNo = '{1}' and RTLot ='{2}' and OrgCode = '{3}'", RecDocNo, MaterialNo, RTLot, LoginData.OrgCode)
                    ds = da.ExecuteDataSet(Sqlstr, "CLIDs")
                    If ds.Tables(0).Rows.Count > CLIDCount Then
                        ReturnBoxID = True
                        CLIDCount = ds.Tables(0).Rows.Count
                    End If
                End If

                Dim myCLIDs = New DataSet
                If ReturnBoxID = True Then
                    Sqlstr = String.Format("Select OrgCode, '' as CLID, Sum(QtyBaseUOM) As QtyBaseUOM, BoxID, MaterialNo, MaterialRevision, BaseUOM, PurOrdNo, PurOrdItem, RecDocNo as RTNo, RTLot, SLOC as SubInv, StorageBin as Locator, StorageType, CONVERT(varchar, ExpDate, 101) as Expdate, Sum(Qty) as Qty, UOM, RecDocItem as TransactionID, PredefinedSubInv, PredefinedLocator, MCPosition, RoHS, StatusCode, ReferenceCLID, StockType, ReasonCode='', IQCStatus = '', Status = '', Message = '' from T_CLMaster with (nolock) " _
                           & "group by OrgCode, BoxID, MaterialNo, MaterialRevision, BaseUOM, PurOrdNo, PurOrdItem, RecDocNo, RTLot, SLOC, StorageBin, StorageType, Expdate,UOM, RecDocItem, PredefinedSubInv, PredefinedLocator, MCPosition, RoHS, StatusCode, ReferenceCLID, StockType having (SLOC IS NULL or SLOC = '') and (StatusCode = 1 or StatusCode = 9 ) and (BoxID IS NOT NULL ) and RecDocNo ='{0}'and MaterialNo = '{1}' and RTLot ='{2}' and OrgCode = '{3}' ", RecDocNo, MaterialNo, RTLot, LoginData.OrgCode)
                Else
                    Sqlstr = String.Format("Select OrgCode, CLID, QtyBaseUOM, BoxID, MaterialNo, MaterialRevision, BaseUOM, PurOrdNo, PurOrdItem, RecDocNo as RTNo, RTLot, SLOC as SubInv, StorageBin as Locator, StorageType, CONVERT(varchar, ExpDate, 101) as Expdate, Qty, UOM, RecDocItem as TransactionID, PredefinedSubInv, PredefinedLocator, MCPosition, RoHS, StatusCode, ReferenceCLID, StockType, ReasonCode='', IQCStatus = '', Status = '', Message = '' from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and (StatusCode = 1 or StatusCode = 9 or (StatusCode = 0 and NOT ReferenceCLID IS NULL)) and RecDocNo ='{0}'and MaterialNo = '{1}' and RTLot ='{2}' and OrgCode = '{3}'", RecDocNo, MaterialNo, RTLot, LoginData.OrgCode)

                    'Slot is available, only read the scanned CLID 
                    If LoginData.ResetFlag = True Then
                        Sqlstr = String.Format("Select OrgCode, CLID, QtyBaseUOM, BoxID, MaterialNo, MaterialRevision, BaseUOM, PurOrdNo, PurOrdItem, RecDocNo as RTNo, RTLot, SLOC as SubInv, StorageBin as Locator, StorageType, CONVERT(varchar, ExpDate, 101) as Expdate, Qty, UOM, RecDocItem as TransactionID, PredefinedSubInv, PredefinedLocator, MCPosition, RoHS, StatusCode, ReferenceCLID, StockType, ReasonCode='', IQCStatus = '', Status = '', Message = '' from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and ( (StatusCode = 1 and CLID ='{0}' ) or (StatusCode = 0 and RecDocNo ='{1}' and MaterialNo = '{2}' and RTLot ='{3}' and ReferenceCLID ='{0}') ) and OrgCode = '{4}'  ", CLID, RecDocNo, MaterialNo, RTLot, LoginData.OrgCode)
                    End If
                End If

                myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDS")

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 OrElse myCLIDs.Tables(0).Rows.Count = 0 Then
                    GetValidSource = Nothing
                    Exit Function
                End If

                If ReturnBoxID = True Then
                    myCLIDs.Tables("CLIDS").Rows(0)("ReasonCode") = CLIDCount.ToString           'return to client for message use

                    Dim DR() As DataRow
                    DR = myCLIDs.Tables("CLIDS").Select(" TransactionID = ' ' or TransactionID IS NULL ")
                    If DR.Length > 0 Then
                        Dim j As Integer
                        For j = 0 To DR.Length - 1
                            Dim MergedID As String
                            Dim LabelID As String = DR(j)("BoxID").ToString

                            Sqlstr = String.Format("Select TOP (1) CLID from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and StatusCode = 1 and (RecDocItem IS NULL) and OrgCode ='{0}' and RTLot = '{1}' and BoxID ='{2}' ", LoginData.OrgCode, RTLot, LabelID)
                            MergedID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                            If MergedID <> "" Then
                                DR(j)("Status") = "N"
                                DR(j)("Message") = "The BoxID contained one MergedID " & MergedID & ", please do Delivery for the MergedID first!"
                            End If
                        Next
                        Return myCLIDs
                    End If
                End If


                'Get IQC Status from Oracle
                Get_IQCStatus(LoginData, myCLIDs)
                If myCLIDs.Tables("CLIDS").Rows(0)("Status") = "N" Then
                    Return myCLIDs
                End If


                'For OSP Raw Parts, the PredefinedSubInv / PredefinedLocator will not be blank, put-away will always deliver materials to those PredefinedSubInv / PredefinedLocator
                If PredefinedSubInv <> "" And PredefinedLocator <> "" Then
                    Return myCLIDs
                End If

                'Slot is available, No need to consider if same material with different Lot sitting in the destination locator
                If LoginData.ResetFlag = True Then
                    Return myCLIDs
                End If

                ds = New DataSet

                'Control message for Random Bin Delivery. Options are: W - Warning message, E - Error message
                Dim CLIDFlag As String
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID004'")
                CLIDFlag = Convert.ToString(da.ExecuteScalar(Sqlstr))

                'For normal Parts, the PredefinedSubInv / PredefinedLocator will be blank
                If Locator <> "" And PredefinedSubInv = "" And PredefinedLocator = "" Then
                    'All Factory use Random Locator, so need to check if same material with different Lot is sitting in the destination locator, but not for OSP Raw material
                    Sqlstr = String.Format("Select TOP (1) MaterialNo, RTLot, SLOC, StorageBin FROM T_CLMaster with (nolock) WHERE  StatusCode = 1 and MaterialNo = '{0}' AND RTLot <> '{1}' AND SLOC = '{2}' AND StorageBin = '{3}' and OrgCode = '{4}'", MaterialNo, RTLot, SubInv, Locator, LoginData.OrgCode)
                    ds = da.ExecuteDataSet(Sqlstr, "DestData")

                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    ElseIf ds.Tables("DestData").Rows.Count > 0 Then
                        'If OrgCode = "353" Or OrgCode = "354" Or OrgCode = "355" Or OrgCode = "365" Then 'Philippine Factory only give warning message
                        '    myCLIDs.Tables("CLIDS").Rows(0)("Status") = "W"
                        'Else
                        '    myCLIDs.Tables("CLIDS").Rows(0)("Status") = "E"           'China Factory give error message
                        'End If

                        'Control message for Random Bin Delivery. Options are: W - Warning message, E - Error message
                        'All Factory will give error message, this control was setup in T_Config  08/28/2012
                        myCLIDs.Tables("CLIDS").Rows(0)("Status") = CLIDFlag
                        myCLIDs.Tables("CLIDS").Rows(0)("Message") = "The same material with different Lot already exists in " + ds.Tables("DestData").Rows(0)("SLOC").ToString + " / " + ds.Tables("DestData").Rows(0)("StorageBin").ToString
                        myCLIDs.AcceptChanges()
                    End If
                End If


                'For NPI parts, the PredefinedSubInv will not be blank, but PredefinedLocator is blank, put-away will always deliver materials to those PredefinedSubInv
                If PredefinedSubInv <> "" And PredefinedLocator = "" Then
                    ds = New DataSet
                    Sqlstr = String.Format("Select DISTINCT MaterialNo, RTLot, SLOC, StorageBin, CLIDFlag='' FROM T_CLMaster with (nolock) WHERE  StatusCode = 1 and MaterialNo = '{0}' AND RTLot <> '{1}' AND SLOC = '{2}' and OrgCode = '{3}'", MaterialNo, RTLot, PredefinedSubInv, LoginData.OrgCode)
                    ds = da.ExecuteDataSet(Sqlstr, "DestData")

                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    ElseIf ds.Tables("DestData").Rows.Count > 0 Then
                        'Record Config Flag for Control message for Random Bin Delivery
                        ds.Tables("DestData").Rows(0)("CLIDFlag") = CLIDFlag
                        myCLIDs.Merge(ds.Tables("DestData"))
                    End If
                End If

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("Putaway-GetValidSource", LoginData.User.ToUpper, "LabelID: " & CLID & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function PutawayPost(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ErrorMsg As String = "RTNo " & Items.Tables("CLIDS").Rows(0)("RTNo") & " for item " & Items.Tables("CLIDS").Rows(0)("MaterialNo") & "; "

            Try
                Dim RTLot, TransactionID As String

                Dim myCLIDs = New DataSet
                myCLIDs = Items.Copy

                Dim DR() As DataRow
                Dim i As Integer

                Dim myDataRow As Data.DataRow
                Dim ItemLists As DataSet = New DataSet
                Dim CLIDS As DataTable = Items.Tables(0).Clone()
                ItemLists.Tables.Add(CLIDS)


                For i = 0 To myCLIDs.Tables("CLIDS").Rows.Count - 1
                    RTLot = myCLIDs.Tables("CLIDS").Rows(i)("RTLot").ToString
                    TransactionID = myCLIDs.Tables("CLIDS").Rows(i)("TransactionID").ToString

                    If TransactionID <> "" Then
                        DR = Nothing
                        DR = ItemLists.Tables("CLIDS").Select(" RTLot = '" & RTLot & "' and TransactionID = '" & TransactionID & "'")
                        If DR.Length = 0 Then
                            myDataRow = ItemLists.Tables("CLIDS").NewRow()
                            myDataRow("MaterialNo") = myCLIDs.Tables("CLIDS").Rows(i)("MaterialNo")
                            myDataRow("RTNo") = myCLIDs.Tables("CLIDS").Rows(i)("RTNo")
                            myDataRow("QtyBaseUOM") = myCLIDs.Tables("CLIDS").Rows(i)("QtyBaseUOM")
                            myDataRow("BaseUOM") = myCLIDs.Tables("CLIDS").Rows(i)("BaseUOM")
                            myDataRow("RTLot") = myCLIDs.Tables("CLIDS").Rows(i)("RTLot")
                            myDataRow("PurOrdNo") = myCLIDs.Tables("CLIDS").Rows(i)("PurOrdNo")
                            myDataRow("PurOrdItem") = myCLIDs.Tables("CLIDS").Rows(i)("PurOrdItem")       'Add PurOrdItem to identify Receipt type is IR or not
                            myDataRow("SubInv") = myCLIDs.Tables("CLIDS").Rows(i)("SubInv")
                            myDataRow("Locator") = myCLIDs.Tables("CLIDS").Rows(i)("Locator")
                            myDataRow("Expdate") = myCLIDs.Tables("CLIDS").Rows(i)("Expdate")
                            myDataRow("Qty") = myCLIDs.Tables("CLIDS").Rows(i)("Qty")
                            myDataRow("UOM") = myCLIDs.Tables("CLIDS").Rows(i)("UOM")
                            myDataRow("TransactionID") = myCLIDs.Tables("CLIDS").Rows(i)("TransactionID")
                            myDataRow("ReasonCode") = myCLIDs.Tables("CLIDS").Rows(i)("ReasonCode").ToString
                            myDataRow("Status") = ""
                            myDataRow("Message") = ""
                            ItemLists.Tables("CLIDS").Rows.Add(myDataRow)
                        Else
                            Dim BaseQty, Qty As Decimal
                            Qty = DR(0)("Qty") + myCLIDs.Tables("CLIDS").Rows(i)("Qty")
                            BaseQty = DR(0)("QtyBaseUOM") + myCLIDs.Tables("CLIDS").Rows(i)("QtyBaseUOM")

                            DR(0)("Qty") = Math.Round(Qty, 5)
                            DR(0)("QtyBaseUOM") = Math.Round(BaseQty, 5)
                            DR(0).AcceptChanges()
                            DR(0).SetAdded()
                        End If
                    End If
                Next

                PutawayPost = New DataSet

                If ItemLists.Tables(0).Rows.Count > 0 Then
                    PutawayPost = Oracle_Putaway(LoginData, ItemLists)
                End If
                If PutawayPost Is Nothing OrElse PutawayPost.Tables.Count = 0 OrElse PutawayPost.Tables(0).Rows.Count = 0 Then
                    Return Nothing
                End If


                DR = Nothing
                DR = PutawayPost.Tables(0).Select("error_message = 'Y'")
                If DR.Length > 0 Then
                    Dim ra As Integer
                    Dim LabelID, SubInv, Locator, ReasonCode, StockType, IQCStatus, StorageType As String

                    For i = 0 To myCLIDs.Tables(0).Rows.Count - 1
                        If myCLIDs.Tables(0).Rows(i)("StatusCode") = "0" Then         'MergedIDs already disabled, no need to do update
                        Else
                            LabelID = myCLIDs.Tables(0).Rows(i)("CLID").ToString
                            SubInv = myCLIDs.Tables(0).Rows(i)("SubInv").ToString
                            Locator = myCLIDs.Tables(0).Rows(i)("Locator").ToString
                            ReasonCode = myCLIDs.Tables(0).Rows(i)("ReasonCode").ToString
                            StockType = myCLIDs.Tables(0).Rows(i)("StockType").ToString.ToUpper
                            IQCStatus = myCLIDs.Tables(0).Rows(i)("IQCStatus").ToString.ToUpper

                            'Slot is available, then Save Slot in filed: StorageType                  -- 06/21/2016
                            StorageType = myCLIDs.Tables(0).Rows(i)("StorageType").ToString

                            'Update StockType from Q/QV to QP, or from S/SV to SP, if IQCStatus = "ACCEPT"  10/23/2014
                            If StockType <> "FTS" AndAlso IQCStatus = "ACCEPT" Then
                                If StockType.Contains("Q") Then StockType = "QP"
                                If StockType.Contains("S") Then StockType = "SP"
                            End If

                            Dim SqlStr As String

                            If LabelID <> "" Then           'CLID
                                SqlStr = String.Format("update T_CLMaster set LastTransaction='Component Delivery', ChangedOn=getdate(), ChangedBy='{0}', SLOC='{1}', StorageBin='{2}', StorageType='{3}', StockType='{4}', ReasonCode='{5}'  where CLID='{6}'", LoginData.User.ToUpper, SubInv, Locator, StorageType, StockType, ReasonCode, LabelID)
                            Else                            'BoxID
                                Dim BoxID As String = myCLIDs.Tables(0).Rows(i)("BoxID").ToString
                                SqlStr = String.Format("update T_CLMaster set LastTransaction='Component Delivery', ChangedOn=getdate(), ChangedBy='{0}', SLOC='{1}', StorageBin='{2}', StorageType='{3}', StockType='{4}', ReasonCode='{5}'  where BoxID='{6}'", LoginData.User.ToUpper, SubInv, Locator, StorageType, StockType, ReasonCode, BoxID)
                            End If
                            ra = da.ExecuteNonQuery(SqlStr)

                            'If LabelID <> "" Then           'CLID
                            '    Dim Clear_BoxID As Boolean
                            '    Clear_BoxID = CleanBoxID(LabelID, LoginData)
                            'End If

                        End If
                    Next
                End If

            Catch ex As Exception
                ErrorLogging("Putaway-PutawayPost", LoginData.User.ToUpper, ErrorMsg & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function Oracle_Putaway(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim BatchID As Integer = p_ds.Tables("CLIDS").Rows(0)("TransactionID")
            Dim ErrorMsg As String = "RTNo " & p_ds.Tables("CLIDS").Rows(0)("RTNo") & " for item " & p_ds.Tables("CLIDS").Rows(0)("MaterialNo") & " and BatchID: " & BatchID

            Dim POStr As String = ""
            If p_ds.Tables("CLIDS").Rows(0)("PurOrdNo").ToString <> "" Then
                POStr = Microsoft.VisualBasic.Mid(p_ds.Tables("CLIDS").Rows(0)("PurOrdNo"), 4, 2)
                If p_ds.Tables("CLIDS").Rows(0)("PurOrdItem").ToString <> "" Then
                    If Not p_ds.Tables("CLIDS").Rows(0)("PurOrdItem").ToString.Contains(".") Then POStr = "IR"
                End If
            End If

            Dim oda As OracleDataAdapter = da.Oda_Insert()

            Try
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_only_deli"
                oda.InsertCommand.Parameters.Add("p_user_id", OracleType.VarChar, 50).Value = LoginData.UserID
                oda.InsertCommand.Parameters.Add("p_batch_id", OracleType.Int32).Value = BatchID
                oda.InsertCommand.Parameters.Add("p_parent_transaction_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_lot_number", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
                oda.InsertCommand.Parameters.Add("p_quantity", OracleType.Double)
                oda.InsertCommand.Parameters.Add("p_unit_of_measure", OracleType.VarChar, 25)
                'oda.InsertCommand.Parameters.Add("p_primary_quantity", OracleType.Double)                 'QtyBaseUOM
                'oda.InsertCommand.Parameters.Add("p_primary_unit_of_measure", OracleType.VarChar, 25)     'BaseUOM
                oda.InsertCommand.Parameters.Add("p_subinventory", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_reason_code", OracleType.VarChar, 500)
                oda.InsertCommand.Parameters.Add("o_return_status", OracleType.VarChar, 10).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_return_message", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput

                oda.InsertCommand.Parameters("p_parent_transaction_id").SourceColumn = "TransactionID"
                oda.InsertCommand.Parameters("p_lot_number").SourceColumn = "RTLot"
                oda.InsertCommand.Parameters("p_lot_expiration_date").SourceColumn = "ExpDate"
                oda.InsertCommand.Parameters("p_quantity").SourceColumn = "Qty"
                oda.InsertCommand.Parameters("p_unit_of_measure").SourceColumn = "UOM"
                'oda.InsertCommand.Parameters("p_primary_quantity").SourceColumn = "QtyBaseUOM"
                'oda.InsertCommand.Parameters("p_primary_unit_of_measure").SourceColumn = "BaseUOM"
                oda.InsertCommand.Parameters("p_subinventory").SourceColumn = "SubInv"
                oda.InsertCommand.Parameters("p_locator").SourceColumn = "Locator"
                oda.InsertCommand.Parameters("p_reason_code").SourceColumn = "ReasonCode"
                oda.InsertCommand.Parameters("o_return_status").SourceColumn = "Status"
                oda.InsertCommand.Parameters("o_return_message").SourceColumn = "Message"

                oda.InsertCommand.Connection.Open()
                oda.Update(p_ds.Tables(0))
                oda.InsertCommand.Connection.Close()
                'oda.Dispose()

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables("CLIDS").Select(" Status = 'N' or Status = ' ' or Status IS Null ")
                If DR.Length = 0 Then
                    'Temporarily Record Delivery records, will delete later   Yudy20100806
                    SaveDelivery(p_ds, BatchID, LoginData.User.ToUpper)

                    Return Putaway_Submit(LoginData, POStr, BatchID, ErrorMsg)
                    Exit Function
                End If

                'Do not delete it which suggested by Charles -- 04/15/2013
                'Delete error from interface table if there has error.  
                'del_delivery_inte(LoginData.UserID, BatchID)

                Dim MSG As DataTable
                Dim myDataRow As Data.DataRow
                Dim myDataColumn As DataColumn
                Dim ErrorList As DataSet = New DataSet

                MSG = New Data.DataTable("MSG")
                myDataColumn = New Data.DataColumn("column_name", System.Type.GetType("System.String"))
                MSG.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("error_message", System.Type.GetType("System.String"))
                MSG.Columns.Add(myDataColumn)

                ErrorList.Tables.Add(MSG)

                'Record and Return Error message
                Dim i As Integer
                Dim ErrMsg As String = ""
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString <> "" Then
                        myDataRow = MSG.NewRow()
                        myDataRow("column_name") = "ERROR"
                        myDataRow("error_message") = DR(i)("Message").ToString
                        MSG.Rows.Add(myDataRow)
                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString
                    End If
                Next
                ErrorLogging("Putaway-Oracle_Putaway1", LoginData.User.ToUpper, ErrorMsg & " with error message: " & ErrMsg, "I")

                Return ErrorList

            Catch oe As Exception
                ErrorLogging("Putaway-Oracle_Putaway", LoginData.User.ToUpper, ErrorMsg & "; " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Putaway_Submit(ByVal LoginData As ERPLogin, ByVal POStr As String, ByVal p_batch_id As Int32, ByVal ErrorMsg As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim Oda As OracleDataAdapter = da.Oda_Sele()

            Dim dsMsg As New DataSet()
            Dim MSG As DataTable
            Dim myDataRow As Data.DataRow
            Dim myDataColumn As DataColumn

            MSG = New Data.DataTable("MSG")
            myDataColumn = New Data.DataColumn("column_name", System.Type.GetType("System.String"))
            MSG.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("error_message", System.Type.GetType("System.String"))
            MSG.Columns.Add(myDataColumn)

            dsMsg.Tables.Add(MSG)

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim Global_RespID As String = 0
                If POStr = "04" Then Global_RespID = GetGlobalRespID(LoginData.Server)

                Dim ds As New DataSet()
                ds.Tables.Add("MSG")
                Dim comm_submit As OracleCommand = da.OraCommand()
                Oda.SelectCommand.CommandType = CommandType.StoredProcedure
                Oda.SelectCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.rcv_submit_req_deliver"
                Oda.SelectCommand.Parameters.Add("o_error_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OrgID  'LoginData.OrgCode
                Oda.SelectCommand.Parameters.Add("p_user_id", OracleType.Int32).Value = LoginData.UserID
                Oda.SelectCommand.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
                Oda.SelectCommand.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID
                Oda.SelectCommand.Parameters.Add("p_global_resp_id", OracleType.Int32).Value = Global_RespID
                Oda.SelectCommand.Parameters.Add("p_batch_id", OracleType.Int32).Value = p_batch_id

                Oda.SelectCommand.Parameters.Add("errbuff", OracleType.VarChar, 4000)
                Oda.SelectCommand.Parameters.Add("retcode", OracleType.VarChar, 240)
                Oda.SelectCommand.Parameters.Add("o_result", OracleType.VarChar, 240)

                Oda.SelectCommand.Parameters("errbuff").Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters("retcode").Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters("o_result").Direction = ParameterDirection.Output

                If Oda.SelectCommand.Connection.State = ConnectionState.Closed Then
                    Oda.SelectCommand.Connection.Open()
                End If
                Oda.Fill(ds, "MSG")
                Oda.SelectCommand.Connection.Close()

                'Record Error message
                Dim i As Integer
                Dim ErrMsg As String = ""
                Dim DR() As DataRow = Nothing
                DR = ds.Tables(0).Select("error_message <> 'Y'")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        ErrMsg = ErrMsg & "; " & DR(i)("column_name").ToString & ": " & DR(i)("error_message").ToString
                    Next
                    ErrorLogging("Putaway-Putaway_Submit1", LoginData.User.ToUpper, ErrorMsg & " with error message " & ErrMsg, "I")

                    'Delete error from interface table after eTrace recorded it in errorlog
                    del_delivery_inte(LoginData.UserID, p_batch_id)
                End If

                Return ds
            Catch oe As Exception
                ErrorLogging("Putaway-Putaway_Submit", LoginData.User.ToUpper, ErrorMsg & ", " & oe.Message & oe.Source, "E")
                myDataRow = MSG.NewRow()
                myDataRow("column_name") = "successful_flag"
                myDataRow("error_message") = "Y"
                MSG.Rows.Add(myDataRow)
                Return dsMsg
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Sub Get_IQCStatus(ByVal LoginData As ERPLogin, ByRef myCLIDs As DataSet)
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i, j As Integer
                Dim DR() As DataRow
                'Dim Sqlstr, MergedID, RTLot As String
                'Dim ds As DataSet = New DataSet

                'DR = Nothing
                'DR = myCLIDs.Tables("CLIDS").Select(" TransactionID = ' ' or TransactionID IS NULL ")
                ''Get Original CLIDs and TransactionID for the MergedIDs if there exists         Yudy 11/27/2009
                'If DR.Length > 0 Then
                '    For i = 0 To DR.Length - 1
                '        RTLot = DR(i)("RTLot").ToString
                '        MergedID = DR(i)("CLID").ToString

                '        Sqlstr = String.Format("Select OrgCode, CLID, QtyBaseUOM, BoxID, MaterialNo, MaterialRevision, BaseUOM, PurOrdNo, PurOrdItem, RecDocNo as RTNo, RTLot, SLOC as SubInv, StorageBin as Locator, CONVERT(varchar, ExpDate, 101) as Expdate, Qty, UOM, RecDocItem as TransactionID, PredefinedSubInv, PredefinedLocator, MCPosition, RoHS, StatusCode, ReferenceCLID, StockType, ReasonCode='', IQCStatus = '', Status = '', Message = '' from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and StatusCode = 0 and OrgCode ='{0}' and RTLot = '{1}' and ReferenceCLID ='{2}' ", LoginData.OrgCode, RTLot, MergedID)
                '        ds = da.ExecuteDataSet(Sqlstr, "CLIDS")

                '        If ds.Tables(0).Rows.Count > 0 Then
                '            myCLIDs.Tables("CLIDS").Merge(ds.Tables(0))
                '        End If
                '        DR(i).AcceptChanges()
                '    Next
                'End If

                Dim TranIDTable As DataTable
                Dim myDataRow As Data.DataRow
                Dim myDataColumn As DataColumn

                TranIDTable = New Data.DataTable("TranIDTable")
                myDataColumn = New Data.DataColumn("TranID", System.Type.GetType("System.String"))
                TranIDTable.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
                TranIDTable.Columns.Add(myDataColumn)

                Dim IQCTable As DataTable
                IQCTable = New Data.DataTable("IQCTable")
                myDataColumn = New Data.DataColumn("IQCStatus", System.Type.GetType("System.String"))
                IQCTable.Columns.Add(myDataColumn)


                Dim IQCStatus As String = ""
                Dim TranIDLists As String = ""

                For i = 0 To myCLIDs.Tables("CLIDS").Rows.Count - 1
                    Dim TransactionID As String = myCLIDs.Tables("CLIDS").Rows(i)("TransactionID").ToString

                    If myCLIDs.Tables("CLIDS").Rows(i)("StockType").ToString = "FTS" Then
                        myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "ACCEPT"
                        myCLIDs.Tables("CLIDS").Rows(i).AcceptChanges()
                    Else
                        If TransactionID <> "" Then
                            DR = Nothing
                            DR = TranIDTable.Select(" TranID = '" & TransactionID & "'")
                            If DR.Length = 0 Then
                                myDataRow = TranIDTable.NewRow()
                                myDataRow("TranID") = TransactionID
                                myDataRow("Status") = ""   'IQCStatus
                                TranIDTable.Rows.Add(myDataRow)

                                TranIDLists = TranIDLists & TransactionID & ","

                            End If
                        End If
                    End If
                    'myCLIDs.AcceptChanges()
                Next


                Dim RTLot As String = ""
                Dim IQCLists As DataSet = New DataSet
                RTLot = myCLIDs.Tables("CLIDS").Rows(0)("RTLot")

                If TranIDLists <> "" Then
                    Try
                        IQCLists = Oracle_IQCStatus(TranIDLists)
                    Catch ex As Exception
                        ErrorLogging("Putaway-Get_IQCStatus", LoginData.User.ToUpper, "RTNo: " & RTLot & ", " & ex.Message & ex.Source, "E")
                    End Try

                    If IQCLists Is Nothing OrElse IQCLists.Tables.Count = 0 OrElse IQCLists.Tables(0).Rows.Count = 0 Then
                        IQCStatus = "NONE"
                    End If

                    For i = 0 To TranIDTable.Rows.Count - 1
                        Dim TransactionID As String = TranIDTable.Rows(i)("TranID").ToString
                        If TranIDTable.Rows(i)("Status").ToString = "" Then
                            If IQCStatus = "NONE" Then
                                TranIDTable.Rows(i)("Status") = IQCStatus
                            ElseIf IQCLists.Tables(0).Rows.Count > 0 Then
                                DR = Nothing
                                DR = IQCLists.Tables(0).Select(" parent_transaction_id = '" & TransactionID & "'")
                                If DR.Length > 0 Then
                                    TranIDTable.Rows(i)("Status") = DR(0)("transaction_type").ToString.ToUpper
                                Else
                                    TranIDTable.Rows(i)("Status") = "NONE"         ' NONE: Standards for Not do IQC yet
                                End If
                            End If
                            TranIDTable.Rows(i).AcceptChanges()
                        End If
                    Next
                    TranIDTable.AcceptChanges()
                End If


                Dim BoxID As String = ""
                For i = 0 To myCLIDs.Tables("CLIDS").Rows.Count - 1
                    If myCLIDs.Tables("CLIDS").Rows(i)("StockType").ToString = "FTS" Then
                    ElseIf myCLIDs.Tables("CLIDS").Rows(i)("StatusCode").ToString = "1" Then
                        Dim TransactionID As String = myCLIDs.Tables("CLIDS").Rows(i)("TransactionID").ToString

                        If TransactionID = "" Then
                            Dim dsCLIDs As DataSet = New DataSet
                            dsCLIDs = myCLIDs.Copy
                            Dim Trow() As DataRow = Nothing

                            Dim CLID As String = myCLIDs.Tables("CLIDS").Rows(i)("CLID").ToString
                            Trow = dsCLIDs.Tables("CLIDS").Select(" ReferenceCLID = '" & CLID & "'")
                            If Trow.Length > 0 Then
                                For j = 0 To Trow.Length - 1
                                    Dim TranID As String = Trow(j)("TransactionID").ToString

                                    DR = Nothing
                                    DR = TranIDTable.Select(" TranID = '" & TranID & "'")
                                    If DR.Length > 0 Then
                                        Dim IQCdr() As DataRow = Nothing
                                        IQCdr = IQCTable.Select(" IQCStatus = '" & DR(0)("Status") & "'")
                                        If IQCdr.Length = 0 Then
                                            If DR(0)("Status") = "ACCEPT" Or DR(0)("Status") = "REJECT" Then
                                                myDataRow = IQCTable.NewRow()
                                                myDataRow("IQCStatus") = DR(0)("Status").ToString
                                                IQCTable.Rows.Add(myDataRow)
                                            End If
                                        End If

                                        myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = DR(0)("Status").ToString
                                        myCLIDs.Tables("CLIDS").Rows(i).AcceptChanges()
                                        If myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "NONE" Then Exit For

                                    End If
                                Next

                                If myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "NONE" Then
                                ElseIf IQCTable.Rows.Count > 1 Then
                                    myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "Both ACCEPT and REJECT is not allowed for IQC"
                                End If
                            End If

                        Else
                            DR = Nothing
                            DR = TranIDTable.Select(" TranID = '" & TransactionID & "'")
                            If DR.Length > 0 Then
                                myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = DR(0)("Status").ToString
                                myCLIDs.Tables("CLIDS").Rows(i).AcceptChanges()
                            End If

                        End If


                        If myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "ACCEPT" Then
                        ElseIf myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "REJECT" Then
                        ElseIf myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus") = "NONE" Then
                            myCLIDs.Tables("CLIDS").Rows(i)("Status") = "N"
                            myCLIDs.Tables("CLIDS").Rows(i)("Message") = "IQC inspection required for this material before delivery "
                        Else
                            myCLIDs.Tables("CLIDS").Rows(i)("Status") = "N"
                            If myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus").ToString <> "" Then
                                'myCLIDs.Tables("CLIDS").Rows(i)("Message") = myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus")
                                myCLIDs.Tables("CLIDS").Rows(i)("Message") = "Invalid IQC Inspection Status: " & myCLIDs.Tables("CLIDS").Rows(i)("IQCStatus")
                            Else
                                myCLIDs.Tables("CLIDS").Rows(i)("Message") = "IQC Inspection Status not found "
                            End If
                        End If

                    End If
                Next
                myCLIDs.AcceptChanges()


            Catch ex As Exception
                ErrorLogging("Putaway-Get_IQCStatus", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try

        End Using

    End Sub

    Public Function Oracle_IQCStatus(ByVal TranIDLists As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("iqc_status")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.get_rcv_iqc_status"

                oda.SelectCommand.Parameters.Add("o_rcv_iqc_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_tran_id_list", OracleType.VarChar, 3000).Value = TranIDLists
                oda.SelectCommand.Connection.Open()

                oda.Fill(ds, "iqc_status")
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                    Return Nothing
                    Exit Function
                End If

                Return ds

            Catch oe As Exception
                ErrorLogging("Putaway-Oracle_IQCStatus", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function del_delivery_inte(ByVal p_user_id As Int32, ByVal p_batch_id As Integer) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()
            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.del_rcv_inte"
                Oda.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
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
                ErrorLogging("Putaway-del_delivery_inte", "", "BatchID: " & p_batch_id & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using

    End Function

    Public Function SaveDelivery(ByVal Items As DataSet, ByVal BatchID As String, ByVal UserName As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i As Integer
                Dim sqlstr As String

                For i = 0 To Items.Tables(0).Rows.Count - 1
                    Dim RTLot, TransactionID, SubInv, Locator, BaseUOM, UOM, Status, Message As String
                    Dim QtyBaseUOM, Qty As Decimal

                    RTLot = Items.Tables(0).Rows(i)("RTLot")
                    TransactionID = Items.Tables(0).Rows(i)("TransactionID")
                    SubInv = Items.Tables(0).Rows(i)("SubInv")
                    Locator = Items.Tables(0).Rows(i)("Locator")
                    BaseUOM = Items.Tables(0).Rows(i)("BaseUOM")
                    UOM = Items.Tables(0).Rows(i)("UOM")
                    Status = Items.Tables(0).Rows(i)("Status")
                    Message = Items.Tables(0).Rows(i)("Message")

                    QtyBaseUOM = IIf(Items.Tables(0).Rows(i)("QtyBaseUOM") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("QtyBaseUOM"))
                    Qty = IIf(Items.Tables(0).Rows(i)("Qty") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Qty"))

                    sqlstr = String.Format("INSERT INTO Z_Delivery ( BatchID, RTLot, TransactionID, SubInv, Locator, QtyBaseUOM, BaseUOM, Qty, UOM, Status, Message, UserName, DateTime ) values ('{0}','{1}','{2}','{3}', '{4}','{5}','{6}','{7}', '{8}', '{9}', '{10}', '{11}', getdate() )", BatchID, RTLot, TransactionID, SubInv, Locator, QtyBaseUOM, BaseUOM, Qty, UOM, Status, Message, UserName)
                    da.ExecuteNonQuery(sqlstr)

                Next
                SaveDelivery = False

            Catch ex As Exception
                ErrorLogging("Putaway-SaveDelivery", UserName, ex.Message & ex.Source, "E")
                SaveDelivery = False
            End Try

        End Using

    End Function

    Public Function GetMRBSubInv(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            GetMRBSubInv = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetMRBSubinv '{0}'", LoginData.OrgID)
                GetMRBSubInv = da.ExecuteDataSet(Sqlstr, "dtMRB")

                Dim ds As DataSet = New DataSet
                ds = Get_ReasonCode(LoginData)
                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                    GetMRBSubInv.Merge(ds.Tables(0))
                End If

                Return GetMRBSubInv
            Catch ex As Exception
                ErrorLogging("Putaway-GetMRBSubInv", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

#End Region


#Region "IQC"

    Public Function ReadBlockDCLN(ByVal ds As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim dsItem As New DataSet

            Dim User As String = ds.Tables(0).Rows(0)("User").ToString
            Dim MaterialNo As String = ds.Tables(0).Rows(0)("MaterialNo").ToString

            Try
                Dim Sqlstr As String
                ds.DataSetName = "dsItem"

                Sqlstr = String.Format("exec sp_ReadBlockDCLN  N'{0}' ", DStoXML(ds))
                dsItem = da.ExecuteDataSet(Sqlstr, "dtItem")

                'For Block Material Movement, first check if there exists records which BLOCK already (means CLID StatusCode = 4 ), 
                'If Yes, delete the records which Status = '' (means CLID StatusCode = 1 ), and only show the BLOCK Items
                If MaterialNo <> "" And MaterialNo <> "DCLNDownload" Then
                    Dim DR() As DataRow
                    DR = dsItem.Tables(0).Select("Status = 'BLOCK' ")
                    If DR.Length > 0 Then
                        Dim DRLine() As DataRow
                        DRLine = dsItem.Tables(0).Select("Status = '' ")
                        If DRLine.Length > 0 Then
                            Dim i As Integer
                            For i = 0 To DRLine.Length - 1
                                DRLine(i).Delete()
                            Next
                            dsItem.Tables(0).AcceptChanges()
                        End If
                    End If
                End If

                Return dsItem

            Catch ex As Exception
                ErrorLogging("IQC-ReadBlockDCLN", User, ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using
    End Function

    Public Function SaveBlockDCLN(ByVal LoginData As ERPLogin, ByVal ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            SaveBlockDCLN = ""
            Dim MaterialNo As String = ds.Tables(0).Rows(0)("MaterialNo").ToString

            Dim BlockType As String = "REC"
            If MaterialNo <> "" Then BlockType = "MMT"

            Try
                Dim Sqlstr As String
                ds.DataSetName = "dsItem"

                Sqlstr = String.Format("exec sp_SaveBlockDCLN '{0}', '{1}', N'{2}',  N'{3}' ", LoginData.OrgCode, BlockType, LoginData.User, DStoXML(ds))
                SaveBlockDCLN = da.ExecuteScalar(Sqlstr).ToString

            Catch ex As Exception
                ErrorLogging("IQC-SaveBlockDCLN", LoginData.User, ex.Message & ex.Source, "E")
                Return "Data update error"
            End Try
        End Using
    End Function

#End Region

End Class


