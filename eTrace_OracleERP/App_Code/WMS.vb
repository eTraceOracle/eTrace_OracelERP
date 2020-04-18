
Imports System
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Xml
Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports eTrace.Data.Repositories
Imports eTrace.Data.Infrastructure
Imports eTrace.Entities
Imports System.Linq
Imports System.Collections.Generic

Public Structure MO_Information
    Public MOList As DataSet
    Public DJ As String
    Public Model As String
    Public DJQty As Double
    Public ProdLine As String
    Public destSubInv As String
    Public destLoc As String
    Public PickedFlag As String
    Public Flag As String
    Public UTurnSubInv As String

    Public dsJob As DataSet

    Public dsCheck_Pick As DataSet

    Public dsAvl As DataSet
    Public Shortage_WithSlot As String
    Public Shortage_NoSlot As String
    Public Shortage_NoOnhand As String

    Public RefQty As Decimal

    Public dsAllCLIDInfo As DataSet
    Public dsProdLine As DataSet
    Public dsLineSlot As DataSet

    Public ErrMsg As String
End Structure


Public Structure MO_List
    Public MOList As DataSet
    Public DJ As String
    Public Model As String
    Public DJQty As Double
    Public ProdLine As String
    Public destSubInv As String
    Public destLoc As String
    Public PickedFlag As String
    Public Flag As String
    Public ErrMsg As String
End Structure

Public Structure SlotShortageList
    Public dsAvl As DataSet
    Public Shortage_WithSlot As String
    Public Shortage_NoSlot As String
    Public Shortage_NoOnhand As String
    Public ErrMsg As String
End Structure

Public Class WMS
    Inherits PublicFunction


#Region "WMS-Allocation"
    Public Function GetActiveEvent(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMSGetActiveEvent '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS-GetActiveEvent", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetActiveJob(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMSGetActiveJob '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS-GetActiveJob", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetActiveEvent_ActiveJob(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                tables = New String() {"ActiveEvent", "ActiveJob"}
                strSQL = New String() {String.Format("exec sp_WMSGetActiveEvent_ActiveJob '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)}
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetActiveEvent_ActiveJob", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetActiveEvent_ActiveJob_LD(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                tables = New String() {"ActiveEvent", "ActiveJob"}
                strSQL = New String() {String.Format("exec sp_WMSGetActiveEvent_ActiveJob_LD '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)}
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetActiveEvent_ActiveJob_LD", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function EventLightOff(ByVal EventID As String, ByVal EventType As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim i As Integer

            Dim strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")
            tables = New String() {"dtSlot"}

            Try
                sql = String.Format("exec sp_WMSEventLightOff '{0}', '{1}', N'{2}'", EventID, EventType, OracleLoginData.User)
                EventLightOff = Convert.ToString(da.ExecuteScalar(sql))

                If EventLightOff = "" Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_LightOffList '{0}', '{1}', N'{2}'", EventID, EventType, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    If EventType = "MO Confirm" Or EventType = "SO Shipment" Or EventType = "Cycle Count" Then
                        If Not dsSlot1 Is Nothing AndAlso dsSlot1.Tables.Count > 0 AndAlso dsSlot1.Tables(0).Rows.Count > 0 Then
                            LEDControlBySlot(dsSlot1, 0, 0)
                        End If
                    Else
                        If Not dsSlot1 Is Nothing AndAlso dsSlot1.Tables.Count > 0 AndAlso dsSlot1.Tables(0).Rows.Count > 0 Then
                            For i = 0 To dsSlot1.Tables(0).Rows.Count - 1
                                LEDControlByRack(dsSlot1.Tables(0).Rows(i)("Rack"), 0)
                            Next
                            'LEDControlByRack("ALL", 0)
                        End If
                    End If

                End If
            Catch ex As Exception
                ErrorLogging("WMS-EventLightOff", OracleLoginData.User, ex.Message & ex.Source, "E")
                EventLightOff = ex.Message & ex.Source
            End Try
        End Using
    End Function

    Public Function GetLocConfig(ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql As String

            Try
                sql = String.Format("exec sp_WMSGetLocConfig N'{0}'", OracleLoginData.User)
                GetLocConfig = Convert.ToString(da.ExecuteScalar(sql))
            Catch ex As Exception
                ErrorLogging("WMS-GetLocConfig", OracleLoginData.User, ex.Message & ex.Source, "E")
                GetLocConfig = "Error: " & ex.Message & ex.Source
            End Try
        End Using
    End Function

    Public Function GetWMSConfig(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String

            Try
                sql = String.Format("exec sp_WMSGetWMSConfig N'{0}'", OracleLoginData.User)
                GetWMSConfig = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("WMS-GetWMSConfig", OracleLoginData.User, ex.Message & ex.Source, "E")
                'GetWMSConfig = Nothing
            End Try
        End Using
    End Function

    Public Function GetItemUsage(ByVal Job As String, ByVal PCBA As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMSGetItemUsage '{0}','{1}', N'{2}'", Job, PCBA, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS-GetItemUsage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetItemUsage_LD(ByVal Job As String, ByVal PCBA As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMSGetItemUsage_LD '{0}','{1}', N'{2}'", Job, PCBA, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS-GetItemUsage_LD", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function WMS_Check_PickedFlag(ByVal Job As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMS_Check_PickedFlag '{0}','{1}', N'{2}'", Job, OracleLoginData.OrgCode, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS_Check_PickedFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
                ds = Nothing
            End Try
            Return ds
        End Using
    End Function

    Public Function Get_MO_Information(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Dim sql, tmpItem As String
            Dim strSQL(), tables(), tables0() As String
            Dim i As Integer
            Dim dr() As DataRow
            Dim ds1 As DataSet
            Dim DmdQty, AvlQty As Decimal

            Try
                ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Start getting MO information from Oracle", "I")

                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_led_mo_info"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_dn_num", OracleType.VarChar, 200).Value = DN
                oda.SelectCommand.Parameters.Add("p_mo_num", OracleType.VarChar, 200).Value = MO
                oda.SelectCommand.Parameters.Add("p_subinv", OracleType.VarChar, 200).Value = SubInv
                oda.SelectCommand.Parameters.Add("p_locator", OracleType.VarChar, 200).Value = Locator
                oda.SelectCommand.Parameters.Add("p_item", OracleType.VarChar, 200).Value = Item

                oda.SelectCommand.Parameters.Add("o_dj_name", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_model", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_start_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_prod_line", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_subinv", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_locator", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_picked_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_onhand_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_exact_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 240).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                Get_MO_Information.Flag = oda.SelectCommand.Parameters("o_success_flag").Value
                Get_MO_Information.ErrMsg = FixNull(oda.SelectCommand.Parameters("o_error_mssg").Value)

                If Get_MO_Information.Flag <> "Y" Then
                    ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Finish getting MO information from Oracle", "I")
                    Exit Function
                End If

                'If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                '    Return Nothing
                '    Exit Function
                'End If

                ds.Tables(0).TableName = "mo_data"
                ds.Tables(1).TableName = "onhand_data"
                ds.Tables(2).TableName = "pickexact_data"

                Get_MO_Information.MOList = ds.Copy
                Get_MO_Information.DJ = FixNull(oda.SelectCommand.Parameters("o_dj_name").Value)
                Get_MO_Information.Model = FixNull(oda.SelectCommand.Parameters("o_dj_model").Value)

                If IsDBNull(oda.SelectCommand.Parameters("o_dj_start_qty").Value) = True Then
                    Get_MO_Information.DJQty = 0
                Else
                    Get_MO_Information.DJQty = oda.SelectCommand.Parameters("o_dj_start_qty").Value
                End If
                Get_MO_Information.ProdLine = FixNull(oda.SelectCommand.Parameters("o_dj_prod_line").Value)
                Get_MO_Information.destSubInv = FixNull(oda.SelectCommand.Parameters("o_dest_subinv").Value)
                Get_MO_Information.destLoc = FixNull(oda.SelectCommand.Parameters("o_dest_locator").Value)
                Get_MO_Information.PickedFlag = FixNull(oda.SelectCommand.Parameters("o_mo_picked_flag").Value)

                ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Finish getting MO information from Oracle", "I")

            Catch oe As Exception
                ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, oe.Message & oe.Source, "E")
                Get_MO_Information.MOList = Nothing
                Get_MO_Information.DJ = ""
                Get_MO_Information.ErrMsg = oe.Message & oe.Source
                Exit Function
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

            ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Start analyzing eTrace records", "I")

            Try
                For i = Get_MO_Information.MOList.Tables("mo_data").Rows.Count - 1 To 0 Step -1
                    If FixNull(Get_MO_Information.MOList.Tables("mo_data").Rows(i)(0)) = "" Then
                        Exit For
                    End If

                    dr = Nothing
                    DmdQty = 0
                    AvlQty = 0

                    If FixNull(Get_MO_Information.MOList.Tables("mo_data").Rows(i)("unallocated_qty")) = "" Then
                        DmdQty = 0
                    Else
                        DmdQty = Get_MO_Information.MOList.Tables("mo_data").Rows(i)("unallocated_qty")
                    End If
                    If FixNull(Get_MO_Information.MOList.Tables("mo_data").Rows(i)("total_available_qty")) = "" Then
                        AvlQty = 0
                    Else
                        AvlQty = Get_MO_Information.MOList.Tables("mo_data").Rows(i)("total_available_qty")
                    End If

                    'dr = Get_MO_Information.MOList.Tables("onhand_data").Select("item_num = '" & Get_MO_Information.MOList.Tables("mo_data").Rows(i)("item_num") & "'")
                    'If DmdQty <= 0 OrElse dr.Length < 1 Then
                    If DmdQty <= 0 OrElse AvlQty <= 0 Then
                        Get_MO_Information.MOList.Tables("mo_data").Rows.RemoveAt(i)
                    End If
                Next
            Catch ex As Exception
                ErrorLogging("WMS-Get_Remove_InvalidRecord", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_Get_JobInform_MO '{0}',N'{1}'", Get_MO_Information.DJ, OracleLoginData.User)
                Get_MO_Information.dsJob = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("WMS-Get_JobInform_MO", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            If Job <> "" Then
                Try
                    sql = String.Format("exec sp_WMS_Check_PickedFlag '{0}','{1}', N'{2}'", Job, OracleLoginData.OrgCode, OracleLoginData.User)

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    Get_MO_Information.dsCheck_Pick = ds1.Copy

                Catch ex As Exception
                    ErrorLogging("WMS_Check_PickedFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try
            End If

            Try
                If (Not Get_MO_Information.MOList.Tables("mo_data") Is Nothing AndAlso Get_MO_Information.MOList.Tables("mo_data").Rows.Count > 0) OrElse (Not Get_MO_Information.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information.MOList.Tables("pickexact_data").Rows.Count > 0) Then
                    'ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, DStoXML(dsMOList), "I")
                    tables0 = New String() {"Avl"}


                    If Not Get_MO_Information.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        Dim dsExact As New DataSet
                        Dim dtExact As New DataTable
                        dtExact = Get_MO_Information.MOList.Tables("pickexact_data").Copy
                        dsExact.Tables.Add(dtExact)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    Else
                        Dim dsMoData As New DataSet
                        Dim dtMoData As New DataTable
                        dtMoData = Get_MO_Information.MOList.Tables("mo_data").Copy
                        dsMoData.Tables.Add(dtMoData)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMoData), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    End If

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    'ds1 = da.ExecuteDataset(strSQL, tables0)

                    Get_MO_Information.dsAvl = ds1.Copy

                    For i = 0 To ds1.Tables(0).Rows.Count - 1
                        If ds1.Tables(0).Rows(i)("SlotQty") >= ds1.Tables(0).Rows(i)("Demand") Then

                        ElseIf ds1.Tables(0).Rows(i)("SlotQty") < ds1.Tables(0).Rows(i)("Demand") Then
                            If ds1.Tables(0).Rows(i)("TtlQty") > ds1.Tables(0).Rows(i)("SlotQty") Then
                                If Get_MO_Information.Shortage_NoSlot = "" Then
                                    Get_MO_Information.Shortage_NoSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information.Shortage_NoSlot = Get_MO_Information.Shortage_NoSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            Else
                                If Get_MO_Information.Shortage_WithSlot = "" Then
                                    Get_MO_Information.Shortage_WithSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information.Shortage_WithSlot = Get_MO_Information.Shortage_WithSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            End If
                        End If
                    Next

                    If Not Get_MO_Information.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        For i = 0 To Get_MO_Information.MOList.Tables("pickexact_data").Rows.Count - 1
                            tmpItem = Get_MO_Information.MOList.Tables("pickexact_data").Rows(i)("item_number")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information.Shortage_NoOnhand = "" Then
                                    Get_MO_Information.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information.Shortage_NoOnhand = Get_MO_Information.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    Else
                        For i = 0 To Get_MO_Information.MOList.Tables("mo_data").Rows.Count - 1
                            tmpItem = Get_MO_Information.MOList.Tables("mo_data").Rows(i)("item_num")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information.Shortage_NoOnhand = "" Then
                                    Get_MO_Information.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information.Shortage_NoOnhand = Get_MO_Information.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    End If
                End If

                Get_MO_Information.ErrMsg = ""

            Catch ex As Exception
                Get_MO_Information.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetRefQty '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information.RefQty = ds1.Tables(0).Rows(0)("Value")

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                'ErrorLogging("GetAllCLIDInfo_LED", OracleLoginData.User, DStoXML(Get_MO_Information.MOList), "I")

                tables = New String() {"RtLot", "CLID"}


                If Not Get_MO_Information.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    Dim dsExact As New DataSet
                    Dim dtExact As New DataTable
                    dtExact = Get_MO_Information.MOList.Tables("pickexact_data").Copy
                    dsExact.Tables.Add(dtExact)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                Else
                    Dim dsOnhand As New DataSet
                    Dim dtOnhand As New DataTable
                    dtOnhand = Get_MO_Information.MOList.Tables("onhand_data").Copy
                    dsOnhand.Tables.Add(dtOnhand)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsOnhand), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                End If
                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information.dsAllCLIDInfo = ds1.Copy
                Get_MO_Information.ErrMsg = ""
            Catch ex As Exception
                Get_MO_Information.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS-GetAllCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetProdLine '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information.dsProdLine = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                tables = New String() {"LineFeeder", "AllLineFeeder"}

                strSQL = New String() {String.Format("exec sp_WMS_GetLineSlot '{0}',N'{1}','{2}','{3}',N'{4}'", OracleLoginData.OrgCode, ProdLine, SubInv, Locator, OracleLoginData.User)}

                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information.dsLineSlot = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetLineSlot", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Get_MO_Information.UTurnSubInv = "ZTU RW1,ZTU RW2"

            ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Finish analyzing eTrace records", "I")
        End Using
    End Function

    Public Function Get_MO_Information_SQL(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Dim sql, tmpItem As String
            Dim strSQL(), tables(), tables0(), tables1() As String
            Dim i As Integer
            Dim dr() As DataRow
            Dim ds1 As DataSet
            Dim DmdQty, AvlQty As Decimal

            Try
                ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, "Start getting MO information from Oracle", "I")

                Dim ds As New DataSet()

                If oda.SelectCommand.Connection.State = ConnectionState.Closed Then
                    oda.SelectCommand.Connection.Open()
                End If

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_led_mo_info"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_dn_num", OracleType.VarChar, 200).Value = DN
                oda.SelectCommand.Parameters.Add("p_mo_num", OracleType.VarChar, 200).Value = MO
                oda.SelectCommand.Parameters.Add("p_subinv", OracleType.VarChar, 200).Value = SubInv
                oda.SelectCommand.Parameters.Add("p_locator", OracleType.VarChar, 200).Value = Locator
                oda.SelectCommand.Parameters.Add("p_item", OracleType.VarChar, 200).Value = Item

                oda.SelectCommand.Parameters.Add("o_dj_name", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_model", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_start_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_prod_line", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_subinv", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_locator", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_picked_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_onhand_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_exact_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 240).Direction = ParameterDirection.Output

                oda.SelectCommand.ExecuteNonQuery()     'Only cancel MO allocation, do not get table records
                oda.SelectCommand.Connection.Close()

                Get_MO_Information_SQL.Flag = oda.SelectCommand.Parameters("o_success_flag").Value
                Get_MO_Information_SQL.ErrMsg = FixNull(oda.SelectCommand.Parameters("o_error_mssg").Value)

                If Get_MO_Information_SQL.Flag <> "Y" Then
                    ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, "Finish getting MO information from Oracle", "I")
                    Exit Function
                End If

                'If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                '    Return Nothing
                '    Exit Function
                'End If

                'Get table records here by Stored Procedure
                tables1 = New String() {"mo_data", "onhand_data", "pickexact_data"}
                Try
                    strSQL = New String() {String.Format("exec ora_get_led_mo_info {0},'{1}', '{2}'", OracleLoginData.OrgID, MO, DN)}
                    ds = da.ExecuteDataset(strSQL, tables1)
                Catch ex As Exception
                    ErrorLogging("WMS_ora_get_led_mo_info", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try

                Get_MO_Information_SQL.MOList = ds.Copy
                Get_MO_Information_SQL.DJ = FixNull(oda.SelectCommand.Parameters("o_dj_name").Value)
                Get_MO_Information_SQL.Model = FixNull(oda.SelectCommand.Parameters("o_dj_model").Value)

                If IsDBNull(oda.SelectCommand.Parameters("o_dj_start_qty").Value) = True Then
                    Get_MO_Information_SQL.DJQty = 0
                Else
                    Get_MO_Information_SQL.DJQty = oda.SelectCommand.Parameters("o_dj_start_qty").Value
                End If
                Get_MO_Information_SQL.ProdLine = FixNull(oda.SelectCommand.Parameters("o_dj_prod_line").Value)
                Get_MO_Information_SQL.destSubInv = FixNull(oda.SelectCommand.Parameters("o_dest_subinv").Value)
                Get_MO_Information_SQL.destLoc = FixNull(oda.SelectCommand.Parameters("o_dest_locator").Value)
                Get_MO_Information_SQL.PickedFlag = FixNull(oda.SelectCommand.Parameters("o_mo_picked_flag").Value)

                ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, "Finish getting MO information from Oracle", "I")

            Catch oe As Exception
                ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, oe.Message & oe.Source, "E")
                Get_MO_Information_SQL.MOList = Nothing
                Get_MO_Information_SQL.DJ = ""
                Get_MO_Information_SQL.ErrMsg = oe.Message & oe.Source
                Exit Function
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

            ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, "Start analyzing eTrace records", "I")

            Try
                For i = Get_MO_Information_SQL.MOList.Tables("mo_data").Rows.Count - 1 To 0 Step -1
                    If FixNull(Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)(0)) = "" Then
                        Exit For
                    End If

                    dr = Nothing
                    DmdQty = 0
                    AvlQty = 0

                    If FixNull(Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("unallocated_qty")) = "" Then
                        DmdQty = 0
                    Else
                        DmdQty = Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("unallocated_qty")
                    End If
                    If FixNull(Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("total_available_qty")) = "" Then
                        AvlQty = 0
                    Else
                        AvlQty = Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("total_available_qty")
                    End If

                    'dr = Get_MO_Information_SQL.MOList.Tables("onhand_data").Select("item_num = '" & Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("item_num") & "'")
                    'If DmdQty <= 0 OrElse dr.Length < 1 Then
                    If DmdQty <= 0 OrElse AvlQty <= 0 Then
                        Get_MO_Information_SQL.MOList.Tables("mo_data").Rows.RemoveAt(i)
                    End If
                Next
            Catch ex As Exception
                ErrorLogging("WMS-Get_Remove_InvalidRecord", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_Get_JobInform_MO '{0}',N'{1}'", Get_MO_Information_SQL.DJ, OracleLoginData.User)
                Get_MO_Information_SQL.dsJob = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("WMS-Get_JobInform_MO", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            If Job <> "" Then
                Try
                    sql = String.Format("exec sp_WMS_Check_PickedFlag '{0}','{1}', N'{2}'", Job, OracleLoginData.OrgCode, OracleLoginData.User)

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    Get_MO_Information_SQL.dsCheck_Pick = ds1.Copy

                Catch ex As Exception
                    ErrorLogging("WMS_Check_PickedFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try
            End If

            Try
                If (Not Get_MO_Information_SQL.MOList.Tables("mo_data") Is Nothing AndAlso Get_MO_Information_SQL.MOList.Tables("mo_data").Rows.Count > 0) OrElse (Not Get_MO_Information_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows.Count > 0) Then
                    'ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, DStoXML(dsMOList), "I")
                    tables0 = New String() {"Avl"}


                    If Not Get_MO_Information_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        Dim dsExact As New DataSet
                        Dim dtExact As New DataTable
                        dtExact = Get_MO_Information_SQL.MOList.Tables("pickexact_data").Copy
                        dsExact.Tables.Add(dtExact)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_SQL.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    Else
                        Dim dsMoData As New DataSet
                        Dim dtMoData As New DataTable
                        dtMoData = Get_MO_Information_SQL.MOList.Tables("mo_data").Copy
                        dsMoData.Tables.Add(dtMoData)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMoData), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_SQL.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    End If

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    'ds1 = da.ExecuteDataset(strSQL, tables0)

                    Get_MO_Information_SQL.dsAvl = ds1.Copy

                    For i = 0 To ds1.Tables(0).Rows.Count - 1
                        If ds1.Tables(0).Rows(i)("SlotQty") >= ds1.Tables(0).Rows(i)("Demand") Then

                        ElseIf ds1.Tables(0).Rows(i)("SlotQty") < ds1.Tables(0).Rows(i)("Demand") Then
                            If ds1.Tables(0).Rows(i)("TtlQty") > ds1.Tables(0).Rows(i)("SlotQty") Then
                                If Get_MO_Information_SQL.Shortage_NoSlot = "" Then
                                    Get_MO_Information_SQL.Shortage_NoSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_SQL.Shortage_NoSlot = Get_MO_Information_SQL.Shortage_NoSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            Else
                                If Get_MO_Information_SQL.Shortage_WithSlot = "" Then
                                    Get_MO_Information_SQL.Shortage_WithSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_SQL.Shortage_WithSlot = Get_MO_Information_SQL.Shortage_WithSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            End If
                        End If
                    Next

                    If Not Get_MO_Information_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        For i = 0 To Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows(i)("item_number")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_SQL.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_SQL.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_SQL.Shortage_NoOnhand = Get_MO_Information_SQL.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    Else
                        For i = 0 To Get_MO_Information_SQL.MOList.Tables("mo_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_SQL.MOList.Tables("mo_data").Rows(i)("item_num")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_SQL.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_SQL.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_SQL.Shortage_NoOnhand = Get_MO_Information_SQL.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    End If
                End If

                Get_MO_Information_SQL.ErrMsg = ""

            Catch ex As Exception
                Get_MO_Information_SQL.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetRefQty '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_SQL.RefQty = ds1.Tables(0).Rows(0)("Value")

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                'ErrorLogging("GetAllCLIDInfo_LED", OracleLoginData.User, DStoXML(Get_MO_Information_SQL.MOList), "I")

                tables = New String() {"RtLot", "CLID"}


                If Not Get_MO_Information_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    Dim dsExact As New DataSet
                    Dim dtExact As New DataTable
                    dtExact = Get_MO_Information_SQL.MOList.Tables("pickexact_data").Copy
                    dsExact.Tables.Add(dtExact)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                Else
                    Dim dsOnhand As New DataSet
                    Dim dtOnhand As New DataTable
                    dtOnhand = Get_MO_Information_SQL.MOList.Tables("onhand_data").Copy
                    dsOnhand.Tables.Add(dtOnhand)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsOnhand), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                End If
                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information_SQL.dsAllCLIDInfo = ds1.Copy
                Get_MO_Information_SQL.ErrMsg = ""
            Catch ex As Exception
                Get_MO_Information_SQL.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS-GetAllCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetProdLine '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_SQL.dsProdLine = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                tables = New String() {"LineFeeder", "AllLineFeeder"}

                strSQL = New String() {String.Format("exec sp_WMS_GetLineSlot '{0}',N'{1}','{2}','{3}',N'{4}'", OracleLoginData.OrgCode, ProdLine, SubInv, Locator, OracleLoginData.User)}

                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information_SQL.dsLineSlot = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetLineSlot", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Get_MO_Information_SQL.UTurnSubInv = "ZTU RW1,ZTU RW2"

            ErrorLogging("WMS-Get_MO_Information_SQL", OracleLoginData.User, "Finish analyzing eTrace records", "I")
        End Using
    End Function

    Public Function Get_MO_Information_LD(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Dim sql, tmpItem As String
            Dim strSQL(), tables(), tables0() As String
            Dim i As Integer
            Dim dr() As DataRow
            Dim ds1 As DataSet
            Dim DmdQty, AvlQty As Decimal

            Try
                ErrorLogging("WMS-Get_MO_Information", OracleLoginData.User, "Start getting MO information from Oracle", "I")

                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_led_mo_info"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_dn_num", OracleType.VarChar, 200).Value = DN
                oda.SelectCommand.Parameters.Add("p_mo_num", OracleType.VarChar, 200).Value = MO
                oda.SelectCommand.Parameters.Add("p_subinv", OracleType.VarChar, 200).Value = SubInv
                oda.SelectCommand.Parameters.Add("p_locator", OracleType.VarChar, 200).Value = Locator
                oda.SelectCommand.Parameters.Add("p_item", OracleType.VarChar, 200).Value = Item

                oda.SelectCommand.Parameters.Add("o_dj_name", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_model", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_start_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_prod_line", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_subinv", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_locator", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_picked_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_onhand_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_exact_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 240).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                Get_MO_Information_LD.Flag = oda.SelectCommand.Parameters("o_success_flag").Value
                Get_MO_Information_LD.ErrMsg = FixNull(oda.SelectCommand.Parameters("o_error_mssg").Value)

                If Get_MO_Information_LD.Flag <> "Y" Then
                    ErrorLogging("WMS-Get_MO_Information_LD", OracleLoginData.User, "Finish getting MO information from Oracle", "I")
                    Exit Function
                End If

                'If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                '    Return Nothing
                '    Exit Function
                'End If

                ds.Tables(0).TableName = "mo_data"
                ds.Tables(1).TableName = "onhand_data"
                ds.Tables(2).TableName = "pickexact_data"

                Get_MO_Information_LD.MOList = ds.Copy
                Get_MO_Information_LD.DJ = FixNull(oda.SelectCommand.Parameters("o_dj_name").Value)
                Get_MO_Information_LD.Model = FixNull(oda.SelectCommand.Parameters("o_dj_model").Value)

                If IsDBNull(oda.SelectCommand.Parameters("o_dj_start_qty").Value) = True Then
                    Get_MO_Information_LD.DJQty = 0
                Else
                    Get_MO_Information_LD.DJQty = oda.SelectCommand.Parameters("o_dj_start_qty").Value
                End If
                Get_MO_Information_LD.ProdLine = FixNull(oda.SelectCommand.Parameters("o_dj_prod_line").Value)
                Get_MO_Information_LD.destSubInv = FixNull(oda.SelectCommand.Parameters("o_dest_subinv").Value)
                Get_MO_Information_LD.destLoc = FixNull(oda.SelectCommand.Parameters("o_dest_locator").Value)
                Get_MO_Information_LD.PickedFlag = FixNull(oda.SelectCommand.Parameters("o_mo_picked_flag").Value)

                ErrorLogging("WMS-Get_MO_Information_LD", OracleLoginData.User, "Finish getting MO information from Oracle", "I")

            Catch oe As Exception
                ErrorLogging("WMS-Get_MO_Information_LD", OracleLoginData.User, oe.Message & oe.Source, "E")
                Get_MO_Information_LD.MOList = Nothing
                Get_MO_Information_LD.DJ = ""
                Get_MO_Information_LD.ErrMsg = oe.Message & oe.Source
                Exit Function
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

            ErrorLogging("WMS-Get_MO_Information_LD", OracleLoginData.User, "Start analyzing eTrace records", "I")

            Try
                For i = Get_MO_Information_LD.MOList.Tables("mo_data").Rows.Count - 1 To 0 Step -1
                    If FixNull(Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)(0)) = "" Then
                        Exit For
                    End If

                    dr = Nothing
                    DmdQty = 0
                    AvlQty = 0

                    If FixNull(Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("unallocated_qty")) = "" Then
                        DmdQty = 0
                    Else
                        DmdQty = Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("unallocated_qty")
                    End If
                    If FixNull(Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("total_available_qty")) = "" Then
                        AvlQty = 0
                    Else
                        AvlQty = Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("total_available_qty")
                    End If

                    'dr = Get_MO_Information_LD.MOList.Tables("onhand_data").Select("item_num = '" & Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("item_num") & "'")
                    'If DmdQty <= 0 OrElse dr.Length < 1 Then
                    If DmdQty <= 0 OrElse AvlQty <= 0 Then
                        Get_MO_Information_LD.MOList.Tables("mo_data").Rows.RemoveAt(i)
                    End If
                Next
            Catch ex As Exception
                ErrorLogging("WMS-Get_Remove_InvalidRecord", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_Get_JobInform_MO '{0}',N'{1}'", Get_MO_Information_LD.DJ, OracleLoginData.User)
                Get_MO_Information_LD.dsJob = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("WMS-Get_JobInform_MO", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            If Job <> "" Then
                Try
                    sql = String.Format("exec sp_WMS_Check_PickedFlag '{0}','{1}', N'{2}'", Job, OracleLoginData.OrgCode, OracleLoginData.User)

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    Get_MO_Information_LD.dsCheck_Pick = ds1.Copy

                Catch ex As Exception
                    ErrorLogging("WMS_Check_PickedFlag", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try
            End If

            Try
                If (Not Get_MO_Information_LD.MOList.Tables("mo_data") Is Nothing AndAlso Get_MO_Information_LD.MOList.Tables("mo_data").Rows.Count > 0) OrElse (Not Get_MO_Information_LD.MOList.Tables("mo_data") Is Nothing AndAlso Get_MO_Information_LD.MOList.Tables("mo_data").Rows.Count > 0) Then
                    'ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, DStoXML(dsMOList), "I")
                    tables0 = New String() {"Avl"}


                    If Not Get_MO_Information_LD.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        Dim dsExact As New DataSet
                        Dim dtExact As New DataTable
                        dtExact = Get_MO_Information_LD.MOList.Tables("pickexact_data").Copy
                        dsExact.Tables.Add(dtExact)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_LD.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    Else
                        Dim dsMoData As New DataSet
                        Dim dtMoData As New DataTable
                        dtMoData = Get_MO_Information_LD.MOList.Tables("mo_data").Copy
                        dsMoData.Tables.Add(dtMoData)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMoData), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_LD.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    End If

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    'ds1 = da.ExecuteDataset(strSQL, tables0)

                    Get_MO_Information_LD.dsAvl = ds1.Copy

                    For i = 0 To ds1.Tables(0).Rows.Count - 1
                        If ds1.Tables(0).Rows(i)("SlotQty") >= ds1.Tables(0).Rows(i)("Demand") Then

                        ElseIf ds1.Tables(0).Rows(i)("SlotQty") < ds1.Tables(0).Rows(i)("Demand") Then
                            If ds1.Tables(0).Rows(i)("TtlQty") > ds1.Tables(0).Rows(i)("SlotQty") Then
                                If Get_MO_Information_LD.Shortage_NoSlot = "" Then
                                    Get_MO_Information_LD.Shortage_NoSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_LD.Shortage_NoSlot = Get_MO_Information_LD.Shortage_NoSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            Else
                                If Get_MO_Information_LD.Shortage_WithSlot = "" Then
                                    Get_MO_Information_LD.Shortage_WithSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_LD.Shortage_WithSlot = Get_MO_Information_LD.Shortage_WithSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            End If
                        End If
                    Next

                    If Not Get_MO_Information_LD.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        For i = 0 To Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows(i)("item_number")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_LD.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_LD.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_LD.Shortage_NoOnhand = Get_MO_Information_LD.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    Else
                        For i = 0 To Get_MO_Information_LD.MOList.Tables("mo_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_LD.MOList.Tables("mo_data").Rows(i)("item_num")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_LD.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_LD.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_LD.Shortage_NoOnhand = Get_MO_Information_LD.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    End If
                End If

                Get_MO_Information_LD.ErrMsg = ""

            Catch ex As Exception
                Get_MO_Information_LD.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetRefQty '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_LD.RefQty = ds1.Tables(0).Rows(0)("Value")

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                'ErrorLogging("GetAllCLIDInfo_LED", OracleLoginData.User, DStoXML(Get_MO_Information_LD.MOList), "I")

                tables = New String() {"RtLot", "CLID"}


                If Not Get_MO_Information_LD.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_LD.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    Dim dsExact As New DataSet
                    Dim dtExact As New DataTable
                    dtExact = Get_MO_Information_LD.MOList.Tables("pickexact_data").Copy
                    dsExact.Tables.Add(dtExact)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                Else
                    Dim dsOnhand As New DataSet
                    Dim dtOnhand As New DataTable
                    dtOnhand = Get_MO_Information_LD.MOList.Tables("onhand_data").Copy
                    dsOnhand.Tables.Add(dtOnhand)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsOnhand), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                End If
                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information_LD.dsAllCLIDInfo = ds1.Copy
                Get_MO_Information_LD.ErrMsg = ""
            Catch ex As Exception
                Get_MO_Information_LD.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS-GetAllCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetProdLine '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_LD.dsProdLine = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            'Try
            '    tables = New String() {"LineFeeder", "AllLineFeeder"}

            '    strSQL = New String() {String.Format("exec sp_WMS_GetLineSlot '{0}',N'{1}','{2}','{3}',N'{4}'", OracleLoginData.OrgCode, ProdLine, SubInv, Locator, OracleLoginData.User)}

            '    ds1 = New DataSet
            '    ds1 = da.ExecuteDataset(strSQL, tables)
            '    Get_MO_Information_LD.dsLineSlot = ds1.Copy

            'Catch ex As Exception
            '    ErrorLogging("WMS_GetLineSlot", OracleLoginData.User, ex.Message & ex.Source, "E")
            'End Try

            ErrorLogging("WMS-Get_MO_Information_LD", OracleLoginData.User, "Finish analyzing eTrace records", "I")
        End Using
    End Function

    Public Function Get_SubinvLoc_for_CS(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim SqlStr As String
            Try
                SqlStr = String.Format("SELECT Value  FROM T_Config with (nolock) WHERE  ConfigID = 'WMS006'")
                Get_SubinvLoc_for_CS = da.ExecuteDataSet(SqlStr, "SubinvLoc")
            Catch ex As Exception
                ErrorLogging("WMS-Get_SubinvLoc_for_CS", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function Get_MO_Information_CS_SQL(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal Job As String, ByVal ProdLine As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String, ByVal TransType As String) As MO_Information
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Dim sql, tmpItem As String
            Dim strSQL(), tables(), tables0(), tables1() As String
            Dim i As Integer
            Dim dr() As DataRow
            Dim ds1 As DataSet
            Dim DmdQty, AvlQty As Decimal

            Try
                ErrorLogging("WMS-Get_MO_Information_CS_SQL", OracleLoginData.User, "Start getting MO information from Oracle", "I")

                Dim ds As New DataSet()

                tables1 = New String() {"mo_data", "onhand_data", "pickexact_data", "parameters"}
                Try
                    strSQL = New String() {String.Format("exec ora_get_led_ldcsmo_info {0},'{1}'", OracleLoginData.OrgID, MO)}
                    ds = da.ExecuteDataset(strSQL, tables1)
                Catch ex As Exception
                    ErrorLogging("Get_MO_Information_CS_SQL", OracleLoginData.User, ex.Message & ex.Source, "E")
                End Try

                Get_MO_Information_CS_SQL.MOList = ds.Copy

                If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso Not ds.Tables(3) Is Nothing AndAlso ds.Tables(3).Rows.Count > 0 Then
                    Get_MO_Information_CS_SQL.DJ = FixNull(ds.Tables(3).Rows(0)("o_dj_name"))
                    Get_MO_Information_CS_SQL.Model = FixNull(ds.Tables(3).Rows(0)("o_dj_model"))

                    If IsDBNull(ds.Tables(3).Rows(0)("o_dj_start_qty")) = True Then
                        Get_MO_Information_CS_SQL.DJQty = 0
                    Else
                        Get_MO_Information_CS_SQL.DJQty = ds.Tables(3).Rows(0)("o_dj_start_qty")
                    End If
                    Get_MO_Information_CS_SQL.ProdLine = FixNull(ds.Tables(3).Rows(0)("o_dj_prod_line"))
                    Get_MO_Information_CS_SQL.destSubInv = FixNull(ds.Tables(3).Rows(0)("o_dest_subinv"))
                    Get_MO_Information_CS_SQL.destLoc = FixNull(ds.Tables(3).Rows(0)("o_dest_locator"))
                    Get_MO_Information_CS_SQL.PickedFlag = FixNull(ds.Tables(3).Rows(0)("o_mo_picked_flag"))

                    Get_MO_Information_CS_SQL.ErrMsg = ""
                Else
                    Get_MO_Information_CS_SQL.ErrMsg = "^WMS-447@"
                End If

                ErrorLogging("WMS-Get_MO_Information_CS_SQL", OracleLoginData.User, "Finish getting MO information from Oracle", "I")

            Catch oe As Exception
                ErrorLogging("WMS-Get_MO_Information_CS_SQL", OracleLoginData.User, oe.Message & oe.Source, "E")
                Get_MO_Information_CS_SQL.MOList = Nothing
                Get_MO_Information_CS_SQL.DJ = ""
                Get_MO_Information_CS_SQL.ErrMsg = oe.Message & oe.Source
                Exit Function
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

            ErrorLogging("WMS-Get_MO_Information_CS_SQL", OracleLoginData.User, "Start analyzing eTrace records", "I")

            Try
                sql = String.Format("exec sp_WMS_Get_JobInform_MO '{0}',N'{1}'", Get_MO_Information_CS_SQL.DJ, OracleLoginData.User)
                Get_MO_Information_CS_SQL.dsJob = da.ExecuteDataSet(sql)
            Catch ex As Exception
                ErrorLogging("WMS-WMS_Get_JobInform_MO", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                If (Not Get_MO_Information_CS_SQL.MOList.Tables("mo_data") Is Nothing AndAlso Get_MO_Information_CS_SQL.MOList.Tables("mo_data").Rows.Count > 0) OrElse (Not Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows.Count > 0) Then
                    'ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, DStoXML(dsMOList), "I")
                    tables0 = New String() {"Avl"}


                    If Not Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        Dim dsExact As New DataSet
                        Dim dtExact As New DataTable
                        dtExact = Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Copy
                        dsExact.Tables.Add(dtExact)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_LD.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    Else
                        Dim dsMoData As New DataSet
                        Dim dtMoData As New DataTable
                        dtMoData = Get_MO_Information_CS_SQL.MOList.Tables("mo_data").Copy
                        dsMoData.Tables.Add(dtMoData)

                        sql = String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMoData), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                        'strSQL = New String() {String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(Get_MO_Information_LD.MOList), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                    End If

                    ds1 = New DataSet
                    ds1 = da.ExecuteDataSet(sql)
                    'ds1 = da.ExecuteDataset(strSQL, tables0)

                    Get_MO_Information_CS_SQL.dsAvl = ds1.Copy

                    For i = 0 To ds1.Tables(0).Rows.Count - 1
                        If ds1.Tables(0).Rows(i)("SlotQty") >= ds1.Tables(0).Rows(i)("Demand") Then

                        ElseIf ds1.Tables(0).Rows(i)("SlotQty") < ds1.Tables(0).Rows(i)("Demand") Then
                            If ds1.Tables(0).Rows(i)("TtlQty") > ds1.Tables(0).Rows(i)("SlotQty") Then
                                If Get_MO_Information_CS_SQL.Shortage_NoSlot = "" Then
                                    Get_MO_Information_CS_SQL.Shortage_NoSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_CS_SQL.Shortage_NoSlot = Get_MO_Information_CS_SQL.Shortage_NoSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            Else
                                If Get_MO_Information_CS_SQL.Shortage_WithSlot = "" Then
                                    Get_MO_Information_CS_SQL.Shortage_WithSlot = ds1.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    Get_MO_Information_CS_SQL.Shortage_WithSlot = Get_MO_Information_CS_SQL.Shortage_WithSlot & ", " & ds1.Tables(0).Rows(i)("MaterialNo")
                                End If
                            End If
                        End If
                    Next

                    If Not Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        For i = 0 To Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows(i)("item_number")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_CS_SQL.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_CS_SQL.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_CS_SQL.Shortage_NoOnhand = Get_MO_Information_CS_SQL.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    Else
                        For i = 0 To Get_MO_Information_CS_SQL.MOList.Tables("mo_data").Rows.Count - 1
                            tmpItem = Get_MO_Information_CS_SQL.MOList.Tables("mo_data").Rows(i)("item_num")
                            dr = Nothing
                            dr = ds1.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If Get_MO_Information_CS_SQL.Shortage_NoOnhand = "" Then
                                    Get_MO_Information_CS_SQL.Shortage_NoOnhand = tmpItem
                                Else
                                    Get_MO_Information_CS_SQL.Shortage_NoOnhand = Get_MO_Information_CS_SQL.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    End If
                End If

                Get_MO_Information_CS_SQL.ErrMsg = ""

            Catch ex As Exception
                Get_MO_Information_CS_SQL.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetRefQty '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_CS_SQL.RefQty = ds1.Tables(0).Rows(0)("Value")

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                'ErrorLogging("GetAllCLIDInfo_LED", OracleLoginData.User, DStoXML(Get_MO_Information_SQL.MOList), "I")

                tables = New String() {"RtLot", "CLID"}


                If Not Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data") Is Nothing AndAlso Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    Dim dsExact As New DataSet
                    Dim dtExact As New DataTable
                    dtExact = Get_MO_Information_CS_SQL.MOList.Tables("pickexact_data").Copy
                    dsExact.Tables.Add(dtExact)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsExact), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                Else
                    Dim dsOnhand As New DataSet
                    Dim dtOnhand As New DataTable
                    dtOnhand = Get_MO_Information_CS_SQL.MOList.Tables("onhand_data").Copy
                    dsOnhand.Tables.Add(dtOnhand)
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsOnhand), SubInv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                End If
                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information_CS_SQL.dsAllCLIDInfo = ds1.Copy
                Get_MO_Information_CS_SQL.ErrMsg = ""
            Catch ex As Exception
                Get_MO_Information_CS_SQL.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS-GetAllCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                sql = String.Format("exec sp_WMS_GetProdLine '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds1 = New DataSet
                ds1 = da.ExecuteDataSet(sql)
                Get_MO_Information_CS_SQL.dsProdLine = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Try
                tables = New String() {"LineFeeder", "AllLineFeeder"}

                strSQL = New String() {String.Format("exec sp_WMS_GetLineSlot '{0}',N'{1}','{2}','{3}',N'{4}'", OracleLoginData.OrgCode, ProdLine, SubInv, Locator, OracleLoginData.User)}

                ds1 = New DataSet
                ds1 = da.ExecuteDataset(strSQL, tables)
                Get_MO_Information_CS_SQL.dsLineSlot = ds1.Copy

            Catch ex As Exception
                ErrorLogging("WMS_GetLineSlot", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try

            Get_MO_Information_CS_SQL.UTurnSubInv = "ZTU RW1,ZTU RW2"

            ErrorLogging("WMS-Get_MO_Information_CS_SQL", OracleLoginData.User, "Finish analyzing eTrace records", "I")
        End Using
    End Function


    Public Function Get_MO_For_LED(ByVal OracleLoginData As ERPLogin, ByVal DN As String, ByVal MO As String, ByVal SubInv As String, ByVal Locator As String, ByVal Item As String) As MO_List
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num

            Try
                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_led_mo_info"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OracleLoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_dn_num", OracleType.VarChar, 200).Value = DN
                oda.SelectCommand.Parameters.Add("p_mo_num", OracleType.VarChar, 200).Value = MO
                oda.SelectCommand.Parameters.Add("p_subinv", OracleType.VarChar, 200).Value = SubInv
                oda.SelectCommand.Parameters.Add("p_locator", OracleType.VarChar, 200).Value = Locator
                oda.SelectCommand.Parameters.Add("p_item", OracleType.VarChar, 200).Value = Item

                oda.SelectCommand.Parameters.Add("o_dj_name", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_model", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_start_qty", OracleType.Double).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dj_prod_line", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_subinv", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_dest_locator", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_picked_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_onhand_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_exact_mo_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 240).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                Get_MO_For_LED.Flag = oda.SelectCommand.Parameters("o_success_flag").Value
                Get_MO_For_LED.ErrMsg = FixNull(oda.SelectCommand.Parameters("o_error_mssg").Value)

                If Get_MO_For_LED.Flag <> "Y" Then
                    Exit Function
                End If

                'If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                '    Return Nothing
                '    Exit Function
                'End If

                ds.Tables(0).TableName = "mo_data"
                ds.Tables(1).TableName = "onhand_data"
                ds.Tables(2).TableName = "pickexact_data"

                Get_MO_For_LED.MOList = ds.Copy
                Get_MO_For_LED.DJ = FixNull(oda.SelectCommand.Parameters("o_dj_name").Value)
                Get_MO_For_LED.Model = FixNull(oda.SelectCommand.Parameters("o_dj_model").Value)

                If IsDBNull(oda.SelectCommand.Parameters("o_dj_start_qty").Value) = True Then
                    Get_MO_For_LED.DJQty = 0
                Else
                    Get_MO_For_LED.DJQty = oda.SelectCommand.Parameters("o_dj_start_qty").Value
                End If
                Get_MO_For_LED.ProdLine = FixNull(oda.SelectCommand.Parameters("o_dj_prod_line").Value)
                Get_MO_For_LED.destSubInv = FixNull(oda.SelectCommand.Parameters("o_dest_subinv").Value)
                Get_MO_For_LED.destLoc = FixNull(oda.SelectCommand.Parameters("o_dest_locator").Value)
                Get_MO_For_LED.PickedFlag = FixNull(oda.SelectCommand.Parameters("o_mo_picked_flag").Value)

            Catch oe As Exception
                ErrorLogging("WMS-Get_MO_For_LDE", OracleLoginData.User, oe.Message & oe.Source, "E")
                Get_MO_For_LED.MOList = Nothing
                Get_MO_For_LED.DJ = ""
                Get_MO_For_LED.ErrMsg = oe.Message & oe.Source
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_RefQty(ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet
            Dim sql As String
            Try
                sql = String.Format("exec sp_WMS_GetRefQty '{0}',N'{1}'", OracleLoginData.OrgCode, OracleLoginData.User)

                ds = New DataSet
                ds = da.ExecuteDataSet(sql)
                Get_RefQty = ds

            Catch ex As Exception
                ErrorLogging("WMS_GetRefQty", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function
    Public Function Get_JobInform_MO(ByVal OracleLoginData As ERPLogin, ByVal DJ As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMS_Get_JobInform_MO '{0}',N'{1}'", DJ, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS-Get_JobInform_MO", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function WMS_CheckSlotShortage(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As SlotShortageList
        Using da As DataAccess = GetDataAccess()
            Dim strSQL, tmpItem As String
            Dim i As Integer
            'Dim WithSlot, NoSlot As Decimal
            Dim ds As DataSet
            Dim dr() As DataRow

            Try
                If Not dsMOList.Tables("mo_data") Is Nothing AndAlso dsMOList.Tables("mo_data").Rows.Count > 0 Then
                    'For i = 0 To dsMOList.Tables("mo_data").Rows.Count - 1
                    '    WithSlot = 0
                    '    NoSlot = 0

                    '    strSQL = String.Format("exec sp_WMS_MatAvlWithSlot '{0}', N'{1}'", dsMOList.Tables("mo_data").Rows(i)("item_num"), OracleLoginData.User)
                    '    WithSlot = Convert.ToString(da.ExecuteScalar(strSQL))

                    '    strSQL = String.Format("exec sp_WMS_MatAvlNotAllSlot '{0}', N'{1}'", dsMOList.Tables("mo_data").Rows(i)("item_num"), OracleLoginData.User)
                    '    WithSlot = Convert.ToString(da.ExecuteScalar(strSQL))

                    'Next

                    'ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, DStoXML(dsMOList), "I")

                    If Not dsMOList.Tables("pickexact_data") Is Nothing AndAlso dsMOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(dsMOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        strSQL = String.Format("exec sp_WMS_MatAvlOnSlot_RTLot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                    Else
                        strSQL = String.Format("exec sp_WMS_MatAvlOnSlot '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                    End If

                    ds = da.ExecuteDataSet(strSQL)
                    WMS_CheckSlotShortage.dsAvl = ds.Copy

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If ds.Tables(0).Rows(i)("SlotQty") >= ds.Tables(0).Rows(i)("Demand") Then

                        ElseIf ds.Tables(0).Rows(i)("SlotQty") < ds.Tables(0).Rows(i)("Demand") Then
                            If ds.Tables(0).Rows(i)("TtlQty") > ds.Tables(0).Rows(i)("SlotQty") Then
                                If WMS_CheckSlotShortage.Shortage_NoSlot = "" Then
                                    WMS_CheckSlotShortage.Shortage_NoSlot = ds.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    WMS_CheckSlotShortage.Shortage_NoSlot = WMS_CheckSlotShortage.Shortage_NoSlot & ", " & ds.Tables(0).Rows(i)("MaterialNo")
                                End If
                            Else
                                If WMS_CheckSlotShortage.Shortage_WithSlot = "" Then
                                    WMS_CheckSlotShortage.Shortage_WithSlot = ds.Tables(0).Rows(i)("MaterialNo")
                                Else
                                    WMS_CheckSlotShortage.Shortage_WithSlot = WMS_CheckSlotShortage.Shortage_WithSlot & ", " & ds.Tables(0).Rows(i)("MaterialNo")
                                End If
                            End If
                        End If
                    Next

                    If Not dsMOList.Tables("pickexact_data") Is Nothing AndAlso dsMOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(dsMOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                        For i = 0 To dsMOList.Tables("pickexact_data").Rows.Count - 1
                            tmpItem = dsMOList.Tables("pickexact_data").Rows(i)("item_number")
                            dr = Nothing
                            dr = ds.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If WMS_CheckSlotShortage.Shortage_NoOnhand = "" Then
                                    WMS_CheckSlotShortage.Shortage_NoOnhand = tmpItem
                                Else
                                    WMS_CheckSlotShortage.Shortage_NoOnhand = WMS_CheckSlotShortage.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    Else
                        For i = 0 To dsMOList.Tables("mo_data").Rows.Count - 1
                            tmpItem = dsMOList.Tables("mo_data").Rows(i)("item_num")
                            dr = Nothing
                            dr = ds.Tables(0).Select("MaterialNo = '" & tmpItem & "'")
                            If dr.Length < 1 Then
                                If WMS_CheckSlotShortage.Shortage_NoOnhand = "" Then
                                    WMS_CheckSlotShortage.Shortage_NoOnhand = tmpItem
                                Else
                                    WMS_CheckSlotShortage.Shortage_NoOnhand = WMS_CheckSlotShortage.Shortage_NoOnhand & ", " & tmpItem
                                End If
                            End If
                        Next
                    End If
                End If

                WMS_CheckSlotShortage.ErrMsg = ""

            Catch ex As Exception
                WMS_CheckSlotShortage.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotShortage", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WMS_CheckSlotAvl(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As SlotShortageList
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim ds As DataSet

            Try
                If Not dsMOList.Tables("pickexact_data") Is Nothing AndAlso dsMOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(dsMOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    strSQL = String.Format("exec sp_WMS_MatSlotAvl_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                Else
                    strSQL = String.Format("exec sp_WMS_MatSlotAvl_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)
                End If

                ds = da.ExecuteDataSet(strSQL)
                WMS_CheckSlotAvl.dsAvl = ds.Copy

                WMS_CheckSlotAvl.ErrMsg = ""

            Catch ex As Exception
                WMS_CheckSlotAvl.ErrMsg = ex.Message.ToString
                ErrorLogging("WMS_CheckSlotAvl", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetAllCLIDInfo_LED(ByVal dsMOList As DataSet, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                ErrorLogging("GetAllCLIDInfo_LED", OracleLoginData.User, DStoXML(dsMOList), "I")

                tables = New String() {"RtLot", "CLID"}

                If Not dsMOList.Tables("pickexact_data") Is Nothing AndAlso dsMOList.Tables("pickexact_data").Rows.Count > 0 AndAlso FixNull(dsMOList.Tables("pickexact_data").Rows(0)("move_order")) <> "" Then
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_Exact'{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                Else
                    strSQL = New String() {String.Format("exec sp_WMSGetAllCLIDInfo_NotExact '{0}','{1}','{2}','{3}',N'{4}'", DStoXML(dsMOList), Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                End If
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetAllCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function
    Public Function GetCLIDInfo_LED(ByVal Component As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                tables = New String() {"RtLot", "CLID"}
                strSQL = New String() {String.Format("exec sp_WMSGetCLIDInfo_LED '{0}', '{1}','{2}','{3}',N'{4}'", Component, Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetCLIDInfo_LED", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetCLIDInfo_LED_ByID(ByVal CLIDList As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                tables = New String() {"RtLot", "CLID"}
                strSQL = New String() {String.Format("exec sp_WMSGetCLIDInfo_LED_ByID '{0}', '{1}',N'{2}'", CLIDList, OracleLoginData.OrgCode, OracleLoginData.User)}
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetCLIDInfo_LED__ByID", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetCLIDInfo_RTLot(ByVal Component As String, ByVal RTLot As String, ByVal Subinv As String, ByVal Locator As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL(), tables() As String
            Dim ds As DataSet
            Try
                tables = New String() {"CLID"}
                strSQL = New String() {String.Format("exec sp_WMSGetCLIDInfo_RTLot '{0}','{1}','{2}','{3}','{4}',N'{5}'", Component, RTLot, Subinv, Locator, OracleLoginData.OrgCode, OracleLoginData.User)}
                ds = da.ExecuteDataset(strSQL, tables)

            Catch ex As Exception
                ErrorLogging("WMS-GetCLIDInfo_RTLot", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return ds
        End Using
    End Function

    Public Function GetCLIDCombination(ByVal srcCLIDs As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim dstCLIDs, tmpCLID, rtCLID As New DataSet
        Dim myDataRow, dstRow As DataRow

        Dim dstCLID As DataTable = New Data.DataTable("dstCLID")
        Dim Count As DataColumn = New DataColumn("Count", System.Type.[GetType]("System.Int32"))
        dstCLID.Columns.Add(Count)
        Dim Sum As DataColumn = New DataColumn("Sum", System.Type.[GetType]("System.Double"))
        dstCLID.Columns.Add(Sum)
        Dim LastNo As DataColumn = New DataColumn("LastNo", System.Type.[GetType]("System.Int32"))
        dstCLID.Columns.Add(LastNo)
        Dim CLIDs As DataColumn = New DataColumn("CLIDs", System.Type.[GetType]("System.String"))
        dstCLID.Columns.Add(CLIDs)
        dstCLIDs.Tables.Add(dstCLID)

        Try

            tmpCLID = dstCLIDs.Clone
            rtCLID = dstCLIDs.Clone

            Dim i, j, k, l, m, n As Integer
            n = srcCLIDs.Tables("CLID").Rows.Count

            For i = 1 To n
                myDataRow = tmpCLID.Tables(0).NewRow
                myDataRow("Count") = 1
                myDataRow("Sum") = srcCLIDs.Tables(0).Rows(i - 1)("Qty")
                myDataRow("LastNo") = i
                myDataRow("CLIDs") = srcCLIDs.Tables(0).Rows(i - 1)("CLID")
                tmpCLID.Tables(0).Rows.Add(myDataRow)

                dstRow = dstCLIDs.Tables(0).NewRow
                dstRow("Count") = 1
                dstRow("Sum") = srcCLIDs.Tables(0).Rows(i - 1)("Qty")
                dstRow("LastNo") = i
                dstRow("CLIDs") = srcCLIDs.Tables(0).Rows(i - 1)("CLID")
                dstCLIDs.Tables(0).Rows.Add(dstRow)
            Next

            For j = 2 To n
                For l = 0 To tmpCLID.Tables(0).Rows.Count
                    If l = tmpCLID.Tables(0).Rows.Count Then
                        tmpCLID.Clear()
                        tmpCLID = rtCLID.Copy
                        rtCLID.Clear()
                        Continue For
                    Else

                        For k = (tmpCLID.Tables(0).Rows(l)("LastNo") + 1) To n
                            myDataRow = rtCLID.Tables(0).NewRow
                            myDataRow("Count") = j
                            myDataRow("Sum") = tmpCLID.Tables(0).Rows(l)("Sum") + srcCLIDs.Tables(0).Rows(k - 1)("Qty")
                            myDataRow("LastNo") = k
                            myDataRow("CLIDs") = tmpCLID.Tables(0).Rows(l)("CLIDs") + "," + srcCLIDs.Tables(0).Rows(k - 1)("CLID")
                            rtCLID.Tables(0).Rows.Add(myDataRow)

                            dstRow = dstCLIDs.Tables(0).NewRow
                            dstRow("Count") = j
                            dstRow("Sum") = tmpCLID.Tables(0).Rows(l)("Sum") + srcCLIDs.Tables(0).Rows(k - 1)("Qty")
                            dstRow("LastNo") = k
                            dstRow("CLIDs") = tmpCLID.Tables(0).Rows(l)("CLIDs") + "," + srcCLIDs.Tables(0).Rows(k - 1)("CLID")
                            dstCLIDs.Tables(0).Rows.Add(dstRow)
                        Next
                    End If
                Next
            Next

            GetCLIDCombination = dstCLIDs.Copy

        Catch ex As Exception
            ErrorLogging("WMS-GetCLIDCombination", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function WMS_Save_Table(ByVal dsCLID As DataSet, ByVal dsItem As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WMS_Save_Table '{0}', '{1}',N'{2}'", DStoXML(dsCLID), DStoXML(dsItem), OracleLoginData.User)
                WMS_Save_Table = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                WMS_Save_Table = ex.Message.ToString
                ErrorLogging("WMS_Save_Table", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WMS_Save_Table2(ByVal dsOracle As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WMS_Save_Table2 '{0}', N'{1}'", DStoXML(dsOracle), OracleLoginData.User)
                WMS_Save_Table2 = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                WMS_Save_Table2 = ex.Message.ToString
                ErrorLogging("WMS_Save_Table2", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WMS_Save_Table3(ByVal dsComb As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_WMS_Save_Table3 '{0}', N'{1}'", DStoXML(dsComb), OracleLoginData.User)
                WMS_Save_Table3 = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                WMS_Save_Table3 = ex.Message.ToString
                ErrorLogging("WMS_Save_Table3", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function WMS_Post_MO_Allocation(ByVal dsMOList As DataSet, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim i As Integer

        Try
            For i = 0 To dsMOList.Tables("OraData").Rows.Count - 1
                If dsMOList.Tables("OraData").Rows(i).RowState = DataRowState.Unchanged Then
                    dsMOList.Tables("OraData").Rows(i).SetAdded()
                End If
            Next

            comm.CommandType = CommandType.StoredProcedure
            comm.CommandText = "apps.xxetr_wip_pkg.create_ledmo_allocation"

            'comm.Parameters.Add("p_org_id", OracleType.Double).Value = 1530
            'comm.Parameters.Add("p_moheader_id", OracleType.Double).Value = 5350019
            'comm.Parameters.Add("p_item_id", OracleType.Double).Value = 117434
            'comm.Parameters.Add("p_subinv", OracleType.VarChar, 240).Value = "ZR0 SM1"
            'comm.Parameters.Add("p_locator_id", OracleType.Double).Value = 108642
            'comm.Parameters.Add("p_lot_num", OracleType.VarChar, 240).Value = "597947"
            'comm.Parameters.Add("p_qty", OracleType.Double).Value = 396
            'comm.Parameters.Add("p_user_id", OracleType.Double).Value = 15889
            'comm.Parameters.Add("p_resp_id", OracleType.Double).Value = 55257
            'comm.Parameters.Add("p_appl_id", OracleType.Double).Value = 401
            'comm.Parameters.Add("o_success_flag", OracleType.VarChar, 10)
            'comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000)

            comm.Parameters.Add("p_org_id", OracleType.Double).SourceColumn = "p_org_id"
            comm.Parameters.Add("p_moheader_id", OracleType.Double).SourceColumn = "p_moheader_id"
            comm.Parameters.Add("p_item_id", OracleType.Double).SourceColumn = "p_item_id"
            comm.Parameters.Add("p_subinv", OracleType.VarChar, 240).SourceColumn = "p_subinv"
            comm.Parameters.Add("p_locator_id", OracleType.Double).SourceColumn = "p_locator_id"
            comm.Parameters.Add("p_lot_num", OracleType.VarChar, 240).SourceColumn = "p_lot_num"
            comm.Parameters.Add("p_qty", OracleType.Double).SourceColumn = "p_qty"
            comm.Parameters.Add("p_user_id", OracleType.Double).SourceColumn = "p_user_id"
            comm.Parameters.Add("p_resp_id", OracleType.Double).SourceColumn = "p_resp_id"
            comm.Parameters.Add("p_appl_id", OracleType.Double).SourceColumn = "p_appl_id"
            comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240).SourceColumn = "o_success_flag"
            comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).SourceColumn = "o_error_mssg"

            comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
            comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

            oda_h.InsertCommand = comm
            oda_h.Update(dsMOList.Tables("OraData"))

            Dim DR() As DataRow = Nothing
            DR = dsMOList.Tables("OraData").Select("o_success_flag <> 'Y'")

            Return dsMOList
        Catch ex As Exception
            dsMOList.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            dsMOList.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging("WMS-Post_MO_Allocation", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return dsMOList
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try

        'Dim da As DataAccess = GetDataAccess()
        'Using connection As New OracleConnection(da._OConnString)
        '    Try
        '        connection.Open()
        '        Dim comm As OracleCommand = connection.CreateCommand()
        '        Dim OrgID As String = GetOrgID(OracleLoginData.OrgCode)

        '        Try
        '            'comm.CommandType = CommandType.StoredProcedure
        '            'comm.CommandText = "apps.XXETR_wip_pkg.initialize"
        '            'comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID) '15904
        '            'comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = OracleLoginData.RespID_WIP   '54050
        '            'comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = OracleLoginData.AppID_WIP
        '            'comm.ExecuteOracleNonQuery("")
        '            'comm.Parameters.Clear()

        '            Dim i As Integer
        '            Dim k As Integer = 1
        '            Dim DR() As DataRow

        '            For i = 0 To dsMOList.Tables(0).Rows.Count - 1
        '                comm.Parameters.Clear()
        '                comm.Transaction = connection.BeginTransaction()
        '                comm.CommandType = CommandType.StoredProcedure
        '                comm.CommandText = "apps.xxetr_wip_pkg.create_ledmo_allocation"

        '                comm.Parameters.Add("p_org_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_org_id")
        '                comm.Parameters.Add("p_moheader_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_moheader_id")
        '                comm.Parameters.Add("p_item_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_item_id")
        '                comm.Parameters.Add("p_subinv", OracleType.VarChar, 240).Value = dsMOList.Tables(0).Rows(i)("p_subinv")
        '                comm.Parameters.Add("p_locator_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_locator_id")
        '                comm.Parameters.Add("p_lot_num", OracleType.VarChar, 240).Value = dsMOList.Tables(0).Rows(i)("p_lot_num")
        '                comm.Parameters.Add("p_qty", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_qty")
        '                comm.Parameters.Add("p_user_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_user_id")
        '                comm.Parameters.Add("p_resp_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_resp_id")
        '                comm.Parameters.Add("p_appl_id", OracleType.Double).Value = dsMOList.Tables(0).Rows(i)("p_appl_id")
        '                comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240).Direction = ParameterDirection.InputOutput
        '                comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput

        '                comm.Parameters("o_success_flag").Value = ""
        '                comm.Parameters("o_error_mssg").Value = ""

        '                comm.ExecuteNonQuery()

        '                dsMOList.Tables(0).Rows(i)("o_success_flag") = comm.Parameters("o_success_flag").Value
        '                dsMOList.Tables(0).Rows(i)("o_error_mssg") = comm.Parameters("o_error_mssg").Value

        '                If dsMOList.Tables(0).Rows(i)("o_success_flag") = "N" Then
        '                    comm.Transaction.Rollback()
        '                    Exit For                                                              'Exit For if there has any error
        '                Else
        '                    comm.Transaction.Commit()
        '                End If
        '            Next
        '            connection.Close()

        '            Return dsMOList

        '        Catch ex As Exception
        '            dsMOList.Tables("transfer_table").Rows(0)("o_return_status") = "N"
        '            dsMOList.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
        '            ErrorLogging("WMS-Post_MO_Allocation", OracleLoginData.User, ex.Message & ex.Source, "E")
        '            Return dsMOList
        '        Finally
        '            If connection.State <> ConnectionState.Closed Then connection.Close()
        '        End Try
        '    Catch ex As Exception
        '        dsMOList.Tables("transfer_table").Rows(0)("o_return_status") = "N"
        '        dsMOList.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
        '        ErrorLogging("WMS-Post_MO_Allocation", OracleLoginData.User, ex.Message & ex.Source, "E")
        '        Return dsMOList
        '    End Try

        'End Using
    End Function

    Public Function WMS_Save_Allocation(ByVal dsHeader As DataSet, ByVal dsItem As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dsSlot1, dsSlot2, dsSlot3 As New DataSet("DS")

            Dim mdr1, mdr2, dr(), dr2(), dr3() As DataRow
            Dim dtSlot1, dtSlot2, dtSlot3 As DataTable
            Dim i, j As Integer

            Try
                strSQL = String.Format("exec sp_WMS_Save_Allocation '{0}', '{1}','{2}',N'{3}'", DStoXML(dsHeader), DStoXML(dsItem), OracleLoginData.OrgCode, OracleLoginData.User)
                WMS_Save_Allocation = Convert.ToString(da.ExecuteScalar(strSQL))

                dsSlot1.DataSetName = "DS"
                dtSlot1 = New DataTable("dtSlot")
                dtSlot1.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                dsSlot1.Tables.Add(dtSlot1)

                dsSlot2.DataSetName = "DS"
                dtSlot2 = New DataTable("dtSlot")
                dtSlot2.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                dsSlot2.Tables.Add(dtSlot2)

                'dtSlot3 = New DataTable("dtSlot")
                'dtSlot3.Columns.Add(New Data.DataColumn("slot", System.Type.GetType("System.String")))
                'dsSlot3.Tables.Add(dtSlot3)

                dr = Nothing
                dr = dsItem.Tables("ItemData").Select("IsPrimary = True")
                If dr.Length > 0 Then
                    For i = 0 To dsItem.Tables("ItemData").Rows.Count - 1
                        If FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot")) <> "" Then
                            If dsItem.Tables("ItemData").Rows(i)("IsPrimary") = True Then
                                'delete un-primary record
                                dr2 = Nothing
                                dr2 = dtSlot2.Select("slot = '" & FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot")) & "'")
                                If dr2.Length > 0 Then
                                    For j = 0 To dr2.Length - 1
                                        dtSlot2.Rows.Remove(dr2(j))
                                    Next
                                End If

                                dr3 = Nothing
                                dr3 = dtSlot1.Select("slot = '" & FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot")) & "'")

                                If dr3.Length > 0 Then

                                Else
                                    mdr1 = dtSlot1.NewRow()
                                    mdr1("slot") = FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot"))
                                    dtSlot1.Rows.Add(mdr1)
                                End If

                            ElseIf dsItem.Tables("ItemData").Rows(i)("IsPrimary") = False Then
                                'skip if it is primary already
                                dr2 = Nothing
                                dr2 = dtSlot1.Select("slot = '" & FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot")) & "'")
                                If dr2.Length > 0 Then

                                Else
                                    dr3 = Nothing
                                    dr3 = dtSlot2.Select("slot = '" & FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot")) & "'")

                                    If dr3.Length > 0 Then

                                    Else
                                        mdr2 = dtSlot2.NewRow()
                                        mdr2("slot") = FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot"))
                                        dtSlot2.Rows.Add(mdr2)
                                    End If
                                End If
                            End If

                            'mdr3 = dtSlot3.NewRow()
                            'mdr3("slot") = FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot"))
                            'dtSlot3.Rows.Add(mdr3)
                        End If
                    Next
                    LEDControlBySlot(dsSlot1, 1, 0)
                    LEDControlBySlot(dsSlot2, 2, 0)
                Else
                    For i = 0 To dsItem.Tables("ItemData").Rows.Count - 1
                        mdr1 = dtSlot1.NewRow()
                        mdr1("slot") = FixNull(dsItem.Tables("ItemData").Rows(i)("InvSlot"))
                        dtSlot1.Rows.Add(mdr1)
                    Next
                    LEDControlBySlot(dsSlot1, 1, 0)
                End If

            Catch ex As Exception
                WMS_Save_Allocation = ex.Message.ToString
                ErrorLogging("WMS_Save_Allocation", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetCycleCountData(ByVal cc_name As String, ByVal seq As Integer, ByVal OracleLoginData As ERPLogin) As GetCycleCount
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

                GetCycleCountData.Flag = oda.SelectCommand.Parameters("o_success_flag").Value.ToString.ToUpper()
                GetCycleCountData.Message = oda.SelectCommand.Parameters("o_error_mssg").Value.ToString

                GetCycleCountData.ds = ds
            Catch ex As Exception
                ErrorLogging("WMS-GetCycleCountData", OracleLoginData.User, ex.Message & ex.Source)
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function

    Public Function GetCycleCountList(ByVal cc_name As String, ByVal OracleLoginData As ERPLogin) As GetCycleCount
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim ds As New DataSet

            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "dbo.ora_get_smtcyclecountlist"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 50).Value = OracleLoginData.OrgCode
                sda.SelectCommand.Parameters.Add("@cc_name", SqlDbType.VarChar, 150).Value = cc_name
                sda.SelectCommand.Parameters.Add("@o_success_flag", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
                sda.SelectCommand.Parameters.Add("@o_error_mssg", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds, "cc_list")
                sda.SelectCommand.Connection.Close()

                GetCycleCountList.Flag = sda.SelectCommand.Parameters("@o_success_flag").Value.ToString.ToUpper()
                GetCycleCountList.Message = sda.SelectCommand.Parameters("@o_error_mssg").Value.ToString

                GetCycleCountList.ds = ds
            Catch ex As Exception
                ErrorLogging("WMS-GetCycleCountRecord", OracleLoginData.User, ex.Message & ex.Source)
                Throw ex
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function

    Public Function PostCycleCountAllocation(ByVal dsCC As DataSet, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim dsList As New DataSet
            Dim strsql1, strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")

            tables = New String() {"dtSlot"}

            Try
                strsql1 = String.Format("exec sp_WMS_Post_CycleCountAllocation N'{0}', '{1}','{2}'", DStoXML(dsCC), OracleLoginData.OrgCode, OracleLoginData.User)
                PostCycleCountAllocation = Convert.ToString(da.ExecuteScalar(strsql1))

                If InStr(PostCycleCountAllocation, "Event ID") > 0 Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_CycleCountAllocation N'{0}', '{1}','{2}'", DStoXML(dsCC), OracleLoginData.OrgCode, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    LEDControlBySlot(dsSlot1, 1, 0)
                End If
            Catch ex As Exception
                PostCycleCountAllocation = ex.Message.ToString
                ErrorLogging("WMS_PostCycleCountAllocation", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function PostCycleCountList(ByVal ccItem As String, ByVal ccRev As String, ByVal ccSub As String, ByVal ccLoc As String, ByVal ccRTLot As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strsql1, strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")

            tables = New String() {"dtSlot"}

            Try

                strsql1 = String.Format("exec sp_WMS_Post_CycleCountList '{0}', '{1}','{2}', '{3}','{4}', '{5}',N'{6}'", ccItem, ccRev, ccSub, ccLoc, ccRTLot, OracleLoginData.OrgCode, OracleLoginData.User)
                PostCycleCountList = Convert.ToString(da.ExecuteScalar(strsql1))

                If InStr(PostCycleCountList, "Event ID") > 0 Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_CycleCountList '{0}', '{1}','{2}', '{3}','{4}', '{5}',N'{6}'", ccItem, ccRev, ccSub, ccLoc, ccRTLot, OracleLoginData.OrgCode, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    LEDControlBySlot(dsSlot1, 1, 0)
                End If
            Catch ex As Exception
                PostCycleCountList = ex.Message.ToString
                ErrorLogging("WMS_PostCycleCountList", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function PostOccupiedAllocation(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strsql1, strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")

            tables = New String() {"dtSlot"}
            Try
                strsql1 = String.Format("exec sp_WMS_Post_OccupiedList '{0}', '{1}',N'{2}'", Rack, OracleLoginData.OrgCode, OracleLoginData.User)
                PostOccupiedAllocation = Convert.ToString(da.ExecuteScalar(strsql1))

                If InStr(PostOccupiedAllocation, "Event ID") > 0 Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_OccupiedList '{0}', '{1}',N'{2}'", Rack, OracleLoginData.OrgCode, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    LEDControlBySlot(dsSlot1, 1, 0)
                End If
            Catch ex As Exception
                PostOccupiedAllocation = ex.Message.ToString
                ErrorLogging("WMS_PostOccupiedAllocation", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function PostEmptyAllocation(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strsql1, strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")

            tables = New String() {"dtSlot"}
            Try
                strsql1 = String.Format("exec sp_WMS_Post_EmptyList '{0}', '{1}',N'{2}'", Rack, OracleLoginData.OrgCode, OracleLoginData.User)
                PostEmptyAllocation = Convert.ToString(da.ExecuteScalar(strsql1))

                If InStr(PostEmptyAllocation, "Event ID") > 0 Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_EmptyList '{0}','{1}',N'{2}'", Rack, OracleLoginData.OrgCode, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    LEDControlBySlot(dsSlot1, 1, 0)
                End If
            Catch ex As Exception
                PostEmptyAllocation = ex.Message.ToString
                ErrorLogging("WMS_PostEmptyAllocation", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function PostConditionalAllocation(ByVal Item As String, ByVal CLID As String, ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim strsql1, strSQL(), tables() As String
            Dim dsSlot1 As New DataSet("DS")

            tables = New String() {"dtSlot"}
            Try
                strsql1 = String.Format("exec sp_WMS_Post_ConditionalList '{0}', '{1}','{2}', '{3}',N'{4}'", Item, CLID, Rack, OracleLoginData.OrgCode, OracleLoginData.User)
                PostConditionalAllocation = Convert.ToString(da.ExecuteScalar(strsql1))

                If InStr(PostConditionalAllocation, "Event ID") > 0 Then
                    strSQL = New String() {String.Format("exec sp_WMS_Get_ConditionalList '{0}', '{1}','{2}', '{3}',N'{4}'", Item, CLID, Rack, OracleLoginData.OrgCode, OracleLoginData.User)}
                    dsSlot1 = da.ExecuteDataset(strSQL, tables)
                    dsSlot1.DataSetName = "DS"

                    LEDControlBySlot(dsSlot1, 1, 0)
                End If
            Catch ex As Exception
                PostConditionalAllocation = ex.Message.ToString
                ErrorLogging("WMS_PostConditionalAllocationn", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function UpdateCLIDMissing(ByVal EventID As String, ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim sql, rst As String
            Try
                sql = String.Format("exec sp_WMS_CLID_UpdateMissing '{0}','{1}', N'{2}'", EventID, CLID, OracleLoginData.User)
                rst = Convert.ToString(da.ExecuteScalar(sql))
            Catch ex As Exception
                rst = ex.Message & ex.Source
                ErrorLogging("WMS-UpdateCLIDMissing", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
            Return rst
        End Using
    End Function

    Public Function WMS_Check_EventID(ByVal EventID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMS_Check_EventID '{0}','{1}', N'{2}'", EventID, OracleLoginData.OrgCode, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS_Check_EventID", OracleLoginData.User, ex.Message & ex.Source, "E")
                ds = Nothing
            End Try
            Return ds
        End Using
    End Function

    Public Function WMS_Check_Rack(ByVal Rack As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sql As String
            Dim ds As DataSet
            Try
                sql = String.Format("exec sp_WMS_Check_RAck '{0}','{1}', N'{2}'", Rack, OracleLoginData.OrgCode, OracleLoginData.User)
                ds = da.ExecuteDataSet(sql)

            Catch ex As Exception
                ErrorLogging("WMS_Check_Rack", OracleLoginData.User, ex.Message & ex.Source, "E")
                ds = Nothing
            End Try
            Return ds
        End Using
    End Function

#End Region

#Region "LED Hardware System Control"
    Dim LEDService1 As LEDService1.WebService = New LEDService1.WebService
    Dim LEDService2 As LEDService2.WebService = New LEDService2.WebService
    Dim LEDService3 As LEDService3.WebService = New LEDService3.WebService
    Dim LEDService4 As LEDService4.WebService = New LEDService4.WebService
    Dim LEDService5 As LEDService5.WebService = New LEDService5.WebService
    Dim LEDService6 As LEDService6.WebService = New LEDService6.WebService
    Dim LEDService7 As LEDService7.WebService = New LEDService7.WebService
    Dim LEDService8 As LEDService8.WebService = New LEDService8.WebService
    Dim LEDService9 As LEDService9.WebService = New LEDService9.WebService
    Dim LEDService10 As LEDService10.WebService = New LEDService10.WebService
    Dim LEDService11 As LEDService11.WebService = New LEDService11.WebService
    Dim LEDService12 As LEDService12.WebService = New LEDService12.WebService
    Dim LEDService13 As LEDService13.WebService = New LEDService13.WebService
    Dim LEDService14 As LEDService14.WebService = New LEDService14.WebService
    Dim LEDService15 As LEDService15.WebService = New LEDService15.WebService
    Dim LEDServiceForLineSide As LEDServiceForLineSide.WebService = New LEDServiceForLineSide.WebService

    'Rack--Rack name(data type String)
    'If Rack = ALL, it Is for all all Racks in Warehouse
    'Code--0/1(Off/On. data type: Integer)
    'return value--True/False
    'Public Function LEDControlByRack1(ByVal RackID As String, ByVal Code As Integer) As Boolean
    '    Using da As DataAccess = GetDataAccess()
    '        Dim strSQL As String
    '        Dim dsControlPC As DataSet
    '        Dim ControlPC As String
    '        Dim StockType As String
    '        Dim i As String = "0"
    '        Try
    '            strSQL = String.Format("exec sp_LEDControlPCofRack '{0}'", RackID)
    '            dsControlPC = da.ExecuteDataSet(strSQL)
    '            For Each dr As DataRow In dsControlPC.Tables(0).Rows
    '                ControlPC = dr("ControlPC").ToString
    '                StockType = dr("StockType").ToString
    '                If StockType = "LineSide" Then
    '                    Return LEDServiceForLineSide.LEDControlByRack(RackID, Code, ControlPC)
    '                Else
    '                    Select Case ControlPC
    '                        Case "CNAPGZHOP141789"
    '                            LEDService1.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOD129721"
    '                            LEDService2.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOD129668"
    '                            LEDService3.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161001"
    '                            LEDService4.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161002"
    '                            LEDService5.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161003"
    '                            LEDService6.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161004"
    '                            LEDService7.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161005"
    '                            LEDService8.LEDControlByRack(RackID, Code)
    '                        Case "CNAPGZHOS161006"
    '                            LEDService9.LEDControlByRack(RackID, Code)
    '                    End Select
    '                End If
    '            Next
    '            Return True
    '        Catch ex As Exception
    '            ErrorLogging("LEDControlByRack", i, ex.Message & ex.Source, "E")
    '            Return False
    '        End Try
    '    End Using
    'End Function

    'Code--0/1/2(Off/On/blink. data type Integer)
    'Interval--time(data type: Integer. 0:Not limit)
    'return value--True/False
    'Public Function LEDControlBySlot1(ByVal slotList As DataSet, ByVal Code As Integer, ByVal Interval As Integer) As Boolean
    '    Using da As DataAccess = GetDataAccess()
    '        Dim strSQL As String
    '        Dim dsControlPC As DataSet
    '        Dim ControlPC As String
    '        Dim StockType As String
    '        Dim i As String = "0"
    '        Try
    '            slotList.Tables(0).Columns(0).ColumnName = "slot"
    '            strSQL = String.Format("exec sp_LEDControlPCofSlot '{0}'", slotList.Tables(0).Rows(0)("slot").ToString)
    '            dsControlPC = da.ExecuteDataSet(strSQL)
    '            ControlPC = dsControlPC.Tables(0).Rows(0)("ControlPC").ToString
    '            StockType = dsControlPC.Tables(0).Rows(0)("StockType").ToString
    '            Interval = Interval * 20
    '            If Code = 0 Then
    '                Interval = 0
    '            End If
    '            If StockType = "LineSide" Then
    '                Return LEDServiceForLineSide.LEDControlBySlot(slotList, Code, Interval, ControlPC)
    '            Else
    '                Select Case ControlPC
    '                    Case "CNAPGZHOP141789"
    '                        i = "1"
    '                        Return LEDService1.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOD129721"
    '                        i = "2"
    '                        Return LEDService2.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOD129668"
    '                        i = "3"
    '                        Return LEDService3.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161001"
    '                        i = "4"
    '                        Return LEDService4.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161002"
    '                        i = "5"
    '                        Return LEDService5.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161003"
    '                        i = "6"
    '                        Return LEDService6.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161004"
    '                        i = "7"
    '                        Return LEDService7.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161005"
    '                        i = "8"
    '                        Return LEDService8.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case "CNAPGZHOS161006"
    '                        i = "9"
    '                        Return LEDService9.LEDControlBySlot(slotList, Code, Interval, StockType)
    '                    Case Else
    '                        i = "10"
    '                        Return False
    '                End Select
    '            End If
    '        Catch ex As Exception
    '            ErrorLogging("LEDControlBySlot", i, ex.Message & ex.Source, "E")
    '            Return False
    '        End Try
    '    End Using
    'End Function


    Public Function LEDControlByRack(ByVal RackID As String, ByVal Code As Integer) As Boolean
        Dim flag = IsXGateRack(RackID)
        Dim oldFlag As Boolean = False, newFlag As Boolean = False
        If Not flag Or String.Equals(RackID, "ALL", StringComparison.OrdinalIgnoreCase) Then
            Using da As DataAccess = GetDataAccess()
                Dim strSQL As String
                Dim dsControlPC As DataSet
                Dim ControlPC As String
                Dim StockType As String
                Dim i As String = "0"
                Try
                    strSQL = String.Format("exec sp_LEDControlPCofRack '{0}'", RackID)
                    dsControlPC = da.ExecuteDataSet(strSQL)
                    For Each dr As DataRow In dsControlPC.Tables(0).Rows
                        ControlPC = dr("ControlPC").ToString
                        StockType = dr("StockType").ToString
                        If StockType = "LineSide" Then
                            Return LEDServiceForLineSide.LEDControlByRack(RackID, Code, ControlPC)
                        Else
                            Select Case ControlPC
                                Case "CNAPGZHOP141789"
                                    LEDService1.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOD129721"
                                    LEDService2.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOD129668"
                                    LEDService3.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161001"
                                    LEDService4.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161002"
                                    LEDService5.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161003"
                                    LEDService6.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161004"
                                    LEDService7.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161005"
                                    LEDService8.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS161006"
                                    LEDService9.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS130003"
                                    LEDService10.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS130016"
                                    LEDService11.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS130004"
                                    LEDService12.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS130955"
                                    LEDService13.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS130011"
                                    LEDService14.LEDControlByRack(RackID, Code)
                                Case "CNAPGZHOS120036"
                                    LEDService15.LEDControlByRack(RackID, Code)
                                Case Else
                                    oldFlag = False
                            End Select
                        End If
                    Next
                    oldFlag = True
                Catch ex As Exception
                    ErrorLogging("LEDControlByRack", i, ex.Message & ex.Source, "E")
                    Return False
                End Try
            End Using
        End If
        If flag Or String.Equals(RackID, "ALL", StringComparison.OrdinalIgnoreCase) Then
            newFlag = LEDControlByNewRack(RackID, Code)
        End If
        Return IIf(flag, newFlag, oldFlag)
    End Function

    Public Function LEDControlBySlot(ByVal slotList As DataSet, ByVal Code As Integer, ByVal Interval As Integer) As Boolean
        Dim oldFlag As Integer = -1, newFlag As Integer = -1

        Dim msg = ";" + Code.ToString() + ";" + Interval.ToString()

        Try
            If (slotList Is Nothing Or slotList.Tables.Count = 0 Or slotList.Tables(0).Rows.Count = 0) Then
                ErrorLogging("LEDControlBySlot", "", "The code is not correct" + DStoXML(slotList) + msg, "I")
                Return False
            End If
            Dim sw = System.Diagnostics.Stopwatch.StartNew()
            slotList.Tables(0).Columns(0).ColumnName = "slot"

            '' In cause the Interval not 0
            If (Code = 0) Then
                Interval = 0
            End If

            ErrorLogging("LEDControlBySlot", slotList.Tables(0).Rows.Count, slotList.GetXml() + ";" + Code.ToString() + ";" + Interval.ToString(), "I")
            Dim slots As List(Of String) = slotList.Tables(0).AsEnumerable().[Select](Function(x) x(0).ToString()).ToList()
            Dim newSlotDS = GetSlots(slots, slotList)
            If slotList IsNot Nothing Then
                Using da As DataAccess = GetDataAccess()
                    Dim strSQL As String
                    Dim dsControlPC As DataSet
                    Dim ControlPC As String
                    Dim StockType As String
                    Dim dsPCandSlot As DataSet
                    Dim i As String = "0"
                    Try
                        slotList.Tables(0).Columns(0).ColumnName = "slot"
                        'strSQL = String.Format("exec sp_LEDControlPCofSlot '{0}'", slotList.Tables(0).Rows(0)("slot").ToString)
                        'dsControlPC = da.ExecuteDataSet(strSQL)
                        'ControlPC = dsControlPC.Tables(0).Rows(0)("ControlPC").ToString
                        'StockType = dsControlPC.Tables(0).Rows(0)("StockType").ToString
                        strSQL = String.Format("exec sp_LEDControlPCofSlots '{0}'", DStoXML(slotList))
                        dsPCandSlot = da.ExecuteDataset(strSQL, 2)
                        If Not (slotList.Tables(0).Columns.Contains("LEDAddr")) Then
                            slotList.Tables(0).Columns.Add("LEDAddr")
                        End If

                        For Each drPC As DataRow In dsPCandSlot.Tables(0).Rows
                            ControlPC = drPC("ControlPC").ToString
                            StockType = drPC("StockType").ToString
                            slotList.Tables(0).Rows.Clear()
                            For Each drSlot As DataRow In dsPCandSlot.Tables(1).Select("ControlPC ='" & ControlPC & "'")
                                Dim dr As DataRow
                                dr = slotList.Tables(0).NewRow
                                dr("slot") = drSlot("slot")
                                dr("LEDAddr") = drSlot("LEDAddr")
                                slotList.Tables(0).Rows.Add(dr)
                            Next
                            slotList.AcceptChanges()
                            Interval = Interval * 20
                            If Code = 0 Then
                                Interval = 0
                            End If
                            If StockType = "LineSide" Then
                                oldFlag = LEDServiceForLineSide.LEDControlBySlot(slotList, Code, Interval, ControlPC)
                            Else
                                Select Case ControlPC
                                    Case "CNAPGZHOP141789"
                                        i = "1"
                                        oldFlag = LEDService1.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOD129721"
                                        i = "2"
                                        oldFlag = LEDService2.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOD129668"
                                        i = "3"
                                        oldFlag = LEDService3.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161001"
                                        i = "4"
                                        oldFlag = LEDService4.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161002"
                                        i = "5"
                                        oldFlag = LEDService5.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161003"
                                        i = "6"
                                        oldFlag = LEDService6.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161004"
                                        i = "7"
                                        oldFlag = LEDService7.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161005"
                                        i = "8"
                                        oldFlag = LEDService8.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS161006"
                                        i = "9"
                                        oldFlag = LEDService9.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS130003"
                                        i = "10"
                                        oldFlag = LEDService10.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS130016"
                                        i = "11"
                                        oldFlag = LEDService11.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS130004"
                                        i = "12"
                                        oldFlag = LEDService12.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS130955"
                                        i = "13"
                                        oldFlag = LEDService13.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS130011"
                                        i = "14"
                                        oldFlag = LEDService14.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case "CNAPGZHOS120036"
                                        i = "15"
                                        oldFlag = LEDService15.LEDControlBySlot(slotList, Code, Interval, StockType)
                                    Case Else
                                        i = "16"
                                        oldFlag = False
                                End Select
                            End If
                        Next

                    Catch ex As Exception
                        ErrorLogging("LEDControlBySlot", i, ex.Message & ex.Source, "E")
                        Return False
                    End Try
                End Using
            End If
            If newSlotDS IsNot Nothing Then
                newFlag = LEDControlByNewSlot(newSlotDS, Code, Interval)
            End If

            sw.[Stop]()
            msg = "Time Consuming(s):" ''+ slotList.Tables(0).Rows.Count.ToString()
            'ErrorLogging("LEDControlBySlot", msg, sw.Elapsed.TotalSeconds.ToString(), "I")

            If oldFlag > 0 And newFlag > 0 Then
                Return True
            ElseIf oldFlag = -1 And newFlag > 0 Then
                Return True
            ElseIf newFlag = -1 And oldFlag > 0 Then
                Return True
            End If
        Catch ex As Exception
            ErrorLogging("LEDControlBySlot", "", ex.Message.ToString(), "E")
        End Try

        Return False
    End Function

#End Region

#Region "LED X-Gate"
    Dim factory = New DbFactory(ConfigurationSettings.AppSettings.Item("eTraceDBConnString"))

    'Rack--Rack name(data type String)
    'If Rack = ALL, it Is for all all Racks in Warehouse
    'Code--0/1(Off/On. data type: Integer)
    'return value--True/False
    Public Function LEDControlByNewRack(ByVal RackID As String, ByVal Code As Integer) As Boolean

        Try
            ErrorLogging("LEDControlByNewRack", Code, RackID, "I")
            If (String.IsNullOrWhiteSpace(RackID)) Then
                ErrorLogging("LEDControlByNewRack", Code, "The RackID can't be empty", "I")
                Return False
            End If

            If (Code > 1) Then
                ErrorLogging("LEDControlByNewRack", RackID, "The code is not correct", "I")
                Return False
            End If

            Dim InvRackRepository = New EntityBaseRepository(Of T_InvRack)(factory)

            If (Not String.Equals(RackID, "ALL", StringComparison.InvariantCultureIgnoreCase)) Then
                Dim invRack = InvRackRepository.AllIncluding(Function(x) x.Rack = RackID, Function(o) o.T_InvSlot).FirstOrDefault()
                If (invRack Is Nothing) Then
                    ErrorLogging("LEDControlByNewRack", Code, "The RackID doesn't existed in T_InvRack", "I")
                    Return False
                End If

                UpdateLightStatus(Code, New List(Of T_InvRack)(New T_InvRack() {invRack}), factory, InvRackRepository)


                Dim teraManager = PtlTeraManager.Instance()
                teraManager.LightByRack(Code, invRack)
            Else



                Dim allRacks = InvRackRepository.AllIncluding(Function(o) o.T_InvSlot).OrderBy(Function(o) o.ControlPC).OrderBy(Function(x) x.Rack).ToList()

                If (allRacks Is Nothing OrElse allRacks.Count = 0) Then
                    ErrorLogging("LEDControlByNewRack", Code, "RackID stil not set up in system", "I")
                    Return False
                End If

                ''UpdateLightStatus(Code, allRacks, factory, InvRackRepository)
                Dim strSQL = String.Format("exec sp_LEDControlByRack '{0}',{1}", RackID, Code)
                Using da As DataAccess = GetDataAccess()
                    da.ExecuteScalar(strSQL)
                End Using
                Dim teraManager = PtlTeraManager.Instance()
                teraManager.LightAllRacks(Code, allRacks)
            End If


        Catch ex As Exception
            Dim er As String = String.Empty ''IIf(String.IsNullOrEmpty(ex.InnerException.ToString()), ex.InnerException, ex.Message.ToString())
            ErrorLogging("LEDControlByNewRack", RackID, ex.Message & ex.Source, "E")
            Return False
        End Try
        Return True
    End Function

    'Code--0/1/2(Off/On/blink. data type Integer)
    'Interval--time(data type: Integer. 0:Not limit)
    'return value--True/False
    Public Function LEDControlByNewSlot(ByVal slotList As DataSet, ByVal Code As Integer, ByVal Interval As Integer) As Boolean
        Dim msg = String.Empty
        Try

            msg = ";" + Code.ToString() + ";" + Interval.ToString()

            If (slotList Is Nothing Or slotList.Tables.Count = 0 Or slotList.Tables(0).Rows.Count = 0) Then
                ErrorLogging("LEDControlByNewSlot", "", "The code is not correct" + DStoXML(slotList) + msg, "I")
                Return False
            End If

            ''ErrorLogging("LEDControlBySlot", slotList.Tables(0).Rows.Count, DStoXML(slotList) + msg, "I")

            Dim slots As List(Of String) = slotList.Tables(0).AsEnumerable().[Select](Function(x) x(0).ToString()).ToList()

            If (Code = 0) Then
                Interval = 0
            End If
            ''For Each item In slots
            ''    msg += item.ToString()
            ''Next
            ''ErrorLogging("LEDControlBySlot", "", "The code is not correct " + msg, "I")

            Dim InvSlotRepository = New EntityBaseRepository(Of T_InvSlot)(factory)
            Dim listOfSlot = InvSlotRepository.FindBy(Function(x) slots.Contains(x.Slot)).OrderBy(Function(o) o.Rack).ThenBy(Function(o) o.LEDAddr).ToList()
            Dim rackIDs = listOfSlot.Select(Function(o) o.Rack).Distinct().ToList()

            Dim InvRackRepository = New EntityBaseRepository(Of T_InvRack)(factory)

            Dim listOfRacks = InvRackRepository.FindBy(Function(o) rackIDs.Contains(o.Rack)).ToList() '


            If (Interval = 0 Or Code = 0) Then
                'UpdateLightStatus(Code, listOfRacks, factory, InvRackRepository, False)  'Rack has been handled on sp_LEDControlBySlot, so comment it.
                ''UpdateSlotStatus(Code, listOfSlot, factory, InvSlotRepository)
                Using da As DataAccess = GetDataAccess()
                    Dim strSQL As String
                    Dim DS As DataSet = New DataSet
                    strSQL = String.Format("exec sp_LEDControlBySlot N'{0}',{1}", slotList.GetXml(), Code)
                    da.ExecuteScalar(strSQL)
                End Using
            End If
            Dim teraManager = PtlTeraManager.Instance()
            teraManager.LightBySlot(Code, listOfRacks, listOfSlot, Interval)

        Catch ex As Exception
            ''Dim er As String = IIf(String.IsNullOrEmpty(ex.InnerException.ToString()), ex.InnerException, ex.Message.ToString())
            ErrorLogging("LEDControlByNewSlot", msg, ex.Message & ex.Source, "E")
            Return False
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Update the T_InvRack and T_InvSlot's status and changedon field. 
    ''' </summary>
    ''' <param name="action"></param>
    ''' <param name="invRacks"></param>
    ''' <param name="factory"></param>
    ''' <param name="dbHelper"></param>
    Private Sub UpdateLightStatus(ByVal action As Integer, ByVal invRacks As List(Of T_InvRack), ByVal factory As DbFactory, ByVal dbHelper As EntityBaseRepository(Of T_InvRack), Optional ByVal isRack As Boolean = True)
        'For Each invRack In invRacks
        '    ''ErrorLogging("UpdateLightStatus", "", invRack.Rack, "E")
        '    invRack.UpdateInvRack(action, isRack)
        '    dbHelper.Edit(invRack)
        'Next
        'Dim _unitOfWork = New UnitOfWork(factory)
        '_unitOfWork.Commit()
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            'Dim DS As DataSet = New DataSet
            Try
                For Each invRack In invRacks
                    strSQL = String.Format("exec sp_LEDControlByRack '{0}',{1}", invRack.Rack, action)
                    da.ExecuteScalar(strSQL)
                Next
            Catch ex As Exception
                ErrorLogging("UpdateLightStatus", "", ex.Message & ex.Source, "E")
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Update the T_InvSlot's status and changedon field.  
    ''' </summary>
    ''' <param name="action"></param>
    ''' <param name="invSlots"></param>
    ''' <param name="factory"></param>
    ''' <param name="dbHelper"></param>
    Private Sub UpdateSlotStatus(ByVal action As Integer, ByVal invSlots As List(Of T_InvSlot), ByVal factory As DbFactory, ByVal dbHelper As EntityBaseRepository(Of T_InvSlot))
        For Each invSlot In invSlots
            invSlot.UpdateInvSlot(action)
            dbHelper.Edit(invSlot)
        Next
        Dim _unitOfWork = New UnitOfWork(factory)
        _unitOfWork.Commit()
        '                strSQL = String.Format("exec sp_LEDControlBySlot N'{0}',{1}", slotList.GetXml(), Code);
        '        ExecuteScalar(strSQL);
        'Using da As DataAccess = GetDataAccess()
        '    Dim strSQL As String
        '    Dim DS As DataSet = New DataSet
        '    Try
        '        For Each invRack In invRacks
        '            strSQL = String.Format("exec sp_LEDControlByRack '{0}',{1}", invRack, action)
        '            DS = da.ExecuteDataSet(strSQL)
        '        Next
        '    Catch ex As Exception
        '        ErrorLogging("UpdateLightStatus", "", ex.Message & ex.Source, "E")
        '    End Try
        'End Using
    End Sub
    ''' <summary>
    ''' Check still led light on, if yes return 1, otherwise, return 0
    ''' </summary>
    ''' <param name="rack"></param>
    ''' <returns></returns>
    Public Function ExistLedLightOnByRack(ByVal rack As String) As Boolean
        Dim InvSlotRepository = New EntityBaseRepository(Of T_InvSlot)(factory)
        Dim invSlots = InvSlotRepository.IsExist(Function(o) o.Rack = rack And o.Status <> 0)
        If (invSlots IsNot Nothing) Then
            ''ErrorLogging("ExistLedLightOnByRack", invSlots.Slot, "E")
            Return True
        End If
        Return False
    End Function
    ''' <summary>
    ''' Check the XGate is connected or not, if yes return 1, otherwise return 0
    ''' </summary>
    ''' <param name="xgateIp"></param>
    ''' <param name="busIndex"></param>
    ''' <returns></returns>
    Public Function XGateIsConnected(ByVal xgateIp As String, ByVal busIndex As Integer) As Boolean
        Try
            Dim teraManager = PtlTeraManager.Instance()
            Return teraManager.CheckXgate(xgateIp, busIndex)
        Catch ex As Exception
            ErrorLogging("IsAvailableXGate", xgateIp + ";" + busIndex.ToString(), ex.Message & ex.Source, "E")
        End Try
        Return False

    End Function

    Public Sub XGateReboot(ByVal xgateIp As String)
        Try
            Dim teraManager = PtlTeraManager.Instance()
            teraManager.Stop()
            teraManager.Reboot(xgateIp)
        Catch ex As Exception
            ErrorLogging("XGateReboot", xgateIp, ex.Message & ex.Source, "E")
        End Try
    End Sub

    Public Function GetAllRacks() As List(Of T_InvRack)

        Dim InvRackRepository = New EntityBaseRepository(Of T_InvRack)(factory)
        Dim allRacks = InvRackRepository.AllIncluding(Function(x) x.LEDAddr.Contains("-"), Function(o) o.T_InvSlot).OrderBy(Function(x) x.Rack).ToList()
        If (allRacks Is Nothing OrElse allRacks.Count = 0) Then
            ErrorLogging("GetAllRacks", "", "RackID stil not set up in system", "I")
            Return Nothing
        End If
        Return allRacks
    End Function


    Private Function IsXGateRack(ByVal rackID As String) As Boolean
        Dim InvRackRepository = New EntityBaseRepository(Of T_InvRack)(factory)
        Dim rack = InvRackRepository.FindBy(Function(o) o.Rack = rackID).FirstOrDefault()
        If rack Is Nothing Then
            ErrorLogging("WS-LEDService1-IsXGateRack", "", "Rack is not exsted " + rackID, "I")
        ElseIf rack.LEDAddr.Contains("-") Then
            Return True
        End If
        Return False
    End Function

    Private Function GetSlots(ByVal slots As List(Of String), ByRef ds As DataSet) As DataSet
        Try
            Dim InvSlotRepository = New EntityBaseRepository(Of T_InvSlot)(factory)
            'Dim listOfSlots = InvSlotRepository.FindBy(Function(x) slots.Contains(x.Slot)).OrderBy(Function(o) o.Rack).ThenBy(Function(o) o.LEDAddr).ToList()
            Dim listOfSlots = InvSlotRepository.FindByNoLock(Function(x) slots.Contains(x.Slot)).OrderBy(Function(o) o.Rack).ThenBy(Function(o) o.LEDAddr).ToList()
            If listOfSlots.Count = 0 Then
                ErrorLogging("WS-LEDService1-GetSlots", "", "Slot not correct", "I")
                ds = Nothing
                Return Nothing
            End If

            Dim oldSlots = listOfSlots.Where(Function(x) x.LEDAddr.Split("-").Length <> 3).ToList()
            Dim newSlots = listOfSlots.Where(Function(x) x.LEDAddr.Split("-").Length = 3).ToList()
            If oldSlots.Count <> 0 Then
                Dim dt As DataTable = ToDataTable(oldSlots)
                dt.TableName = "dtSlot"
                dt.Columns(0).ColumnName = "slot"
                Dim tmpDs As New DataSet("DS")
                tmpDs.Tables.Add(dt)
                ds = tmpDs
            Else
                ds = Nothing
            End If

            If newSlots.Count <> 0 Then
                Dim dt As DataTable = ToDataTable(newSlots)
                dt.TableName = "dtSlot"
                dt.Columns(0).ColumnName = "slot"
                Dim tmpDs As New DataSet("DS")
                tmpDs.Tables.Add(dt)
                Return tmpDs
            End If
        Catch ex As Exception
            ds = Nothing
            ErrorLogging("WS-LEDService1-GetSlots", "", ex.Message.ToString(), "E")
        End Try

        Return Nothing
    End Function


    Public Shared Function ToDataTable(Of T)(ByVal items As List(Of T)) As DataTable
        Dim dataTable As New DataTable(GetType(T).Name)

        'Get all the properties
        Dim Props As System.Reflection.PropertyInfo() = GetType(T).GetProperties(System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.Instance)
        For Each prop As System.Reflection.PropertyInfo In Props
            'Defining type of data column gives proper data table 
            Dim type = (If(prop.PropertyType.IsGenericType AndAlso prop.PropertyType.GetGenericTypeDefinition() = GetType(Nullable(Of )), Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType))
            'Setting column names as Property names
            dataTable.Columns.Add(prop.Name, type)
        Next
        For Each item As T In items
            Dim values = New Object(Props.Length - 1) {}
            For i As Integer = 0 To Props.Length - 1
                'inserting property values to datatable rows
                values(i) = Props(i).GetValue(item, Nothing)
            Next
            dataTable.Rows.Add(values)
        Next
        'put a breakpoint here and check datatable
        Return dataTable
    End Function

#End Region

End Class


Module EntitiesExtensions

    <Runtime.CompilerServices.Extension()>
    Public Sub UpdateInvRack(ByVal invRack As T_InvRack, ByVal code As Integer, ByVal isRack As Boolean)

        If (code = 2) Then
            code = 1
        End If
        invRack.Status = code.ToString()
        invRack.ChangedOn = DateTime.Now

        If isRack Then
            For Each item In invRack.T_InvSlot
                item.Status = code.ToString()
                item.ChangedOn = DateTime.Now
            Next
        End If


    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub UpdateInvSlot(ByVal invSlot As T_InvSlot, ByVal code As Integer)
        invSlot.Status = code.ToString()
        invSlot.ChangedOn = DateTime.Now
    End Sub
End Module

