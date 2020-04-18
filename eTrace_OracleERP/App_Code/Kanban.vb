Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading

Public Structure KBdata
    Public kbID As String
    Public strItem As String
    Public strSubInv As String
    Public strQty As String
    Public strLoc As String


End Structure


Public Class Kanban


    Inherits PublicFunction

#Region "Change CLID Expired Date"
    Public Function ChangeExpdateUser(ByVal UserName As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim vCount As Integer
            Dim Sqlstr As String
            Dim ReturnValue As Boolean
            Try
                Sqlstr = String.Format(" select count(1) vCount from T_Config where ConfigID = 'EXPUSER' and Value like '%{0}%' ", UserName.ToUpper)
                vCount = Convert.ToDouble(da.ExecuteScalar(Sqlstr))
                If vCount = 0 Then
                    ReturnValue = False
                Else
                    ReturnValue = True
                End If

                Return ReturnValue
            Catch oe As Exception
                ErrorLogging("ChangeExpdateUser", "", oe.Message & oe.Source, "E")
                Return False
            End Try
        End Using
    End Function
    Public Function BatchChangeExpdate(ByVal BatchID As String, ByVal OracleERPLogin As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim ErrMsg As String
            sda = da.Sda_Sele()
            Try
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_BatchExpDate_Update"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@BatchID", SqlDbType.VarChar, 150)
                sda.SelectCommand.Parameters.Add("@User", SqlDbType.VarChar, 100)
                sda.SelectCommand.Parameters.Add("@ErrMsg", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
                sda.SelectCommand.Parameters("@BatchID").Value = BatchID
                sda.SelectCommand.Parameters("@User").Value = OracleERPLogin.User.ToUpper
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                ErrMsg = sda.SelectCommand.Parameters("@ErrMsg").Value.ToString()
                If ErrMsg <> "" Then
                    Return False
                Else
                    Return True
                End If
            Catch oe As Exception
                ErrorLogging("BatchChangeExpdate", "", oe.Message & oe.Source, "E")
                Return False
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function CheckCLIDExpFormat(ByVal BatchID As String, ByVal mydsCLID As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            sda = da.Sda_Sele()
            Dim i As Integer
            Dim ds As New DataSet
            Dim strSql As String
            Try
                strSql = String.Format("Delete T_BatchChangeExpDate where BatchID like '{0}%'", Left(BatchID, BatchID.Length - 14))
                da.ExecuteNonQuery(strSql)
                For i = 0 To mydsCLID.Tables(0).Rows.Count - 1
                    Dim table1 As DataRow = mydsCLID.Tables(0).Rows(i)
                    strSql = String.Format("INSERT INTO T_BatchChangeExpDate ([BatchID],[OrgCode],[CLID],[Item],[Subinventory],[Locator],[OldExpDate],[NewExpDate])  values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", BatchID, OracleERPLogin.OrgCode, table1("CLID"), table1("Item"), table1("Subinventory"), table1("Locator"), table1("OldExpDate"), table1("NewExpDate"))
                    da.ExecuteNonQuery(strSql)
                Next
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_BatchExpDate_CLIDValid"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@BatchID", SqlDbType.VarChar, 150)
                sda.SelectCommand.Parameters("@BatchID").Value = BatchID
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
                Return ds
            Catch oe As Exception
                ErrorLogging("CheckCLIDExpFormat", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function
#End Region
#Region "UploadEJITIPP"
    Public Function getProdFloor(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow
            Dim result As DataSet
            Try
                strsql = String.Format("select DISTINCT ProdFloor from T_eJITIPPHeader order by ProdFloor")
                getProdFloor = da.ExecuteDataSet(strsql)
            Catch ex As Exception
                ErrorLogging("UploadEJITIPP-getProdFloor", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function
    Public Function getEjitIPPList(ByVal startDate As String, ByVal endDate As String, ByVal productionfloor As String, ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow
            Dim result As DataSet
            Try
                strsql = String.Format("select DISTINCT IPPID,CreatedOn,CreatedBy from T_eJITIPPHeader where  CreatedOn between  '{0}'and '{1}'  and ProdFloor='{2}' and OrgCode='{3}' ", startDate, endDate, productionfloor, LoginData.OrgCode)
                result = da.ExecuteDataSet(strsql)
                If (result Is Nothing) Then
                    myDataRow = p_ds.Tables("IPP").NewRow()
                    myDataRow("IPPID") = -1
                    myDataRow("CreatedOn") = "No record"
                    p_ds.Tables(0).Rows.Add(myDataRow)
                Else
                    myDataRow = p_ds.Tables("IPP").NewRow()
                    myDataRow("IPPID") = -1
                    myDataRow("CreatedOn") = "Please select..."
                    p_ds.Tables(0).Rows.Add(myDataRow)
                    For Each dr As DataRow In result.Tables(0).Rows
                        myDataRow = p_ds.Tables("IPP").NewRow()
                        myDataRow("IPPID") = dr(0).ToString
                        myDataRow("CreatedOn") = dr(1).ToString & "   " & dr(2).ToString
                        p_ds.Tables("IPP").Rows.Add(myDataRow)
                    Next
                End If
                getEjitIPPList = p_ds
            Catch ex As Exception
                ErrorLogging("UploadEJITIPP-getEjitIPPList", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function
    Public Function getIPPDetail(ByVal IPPId As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow
            Dim result As DataTable
            getIPPDetail = New DataSet
            Try
                strsql = String.Format("select IPPID,ProdFloor,CreatedOn,CreatedBy from T_eJITIPPHeader where IPPID= {0}", IPPId)
                result = da.ExecuteDataTable(strsql)
                getIPPDetail.Tables.Add(result)
                strsql = "select DJ,Assembly,IPPQty,IPPDate,[Group] from T_eJITIPPItem where IPPID= {0}"
                strsql = String.Format(strsql, IPPId)
                result = da.ExecuteDataTable(strsql)
                getIPPDetail.Tables.Add(result)
            Catch ex As Exception
                ErrorLogging("UploadEJITIPP-getIPPDetail", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function
    Public Function Upload_EJITIPP(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As Boolean
        Upload_EJITIPP = False
        Using da As DataAccess = GetDataAccess()
            Dim ra As Integer
            Dim strHeader, strDetail, strsql As String
            Dim i As Integer
            Dim result1 As New Object
            Dim KBQty As Decimal = 0
            Dim MPQ As Decimal = 0
            Dim OnhandQty As Decimal = 0
            Dim mcqty As Decimal = 0
            Dim WipQty As Decimal = 0
            Dim Shippinginstruction As String
            Dim IppId As String
            Try
                strsql = String.Format("SELECT NEXT VALUE FOR SeQ_EJITIPPID")
                IppId = Convert.ToString(da.ExecuteScalar(strsql))
                For i = 0 To p_ds.Tables(0).Rows.Count - 1
                    Dim table1 As DataRow = p_ds.Tables(0).Rows(i)
                    If i = 0 Then
                        ChangeEjitIpp(table1("ProdFloor"))
                        strHeader = String.Format("INSERT INTO T_eJITIPPHeader (IPPID,OrgCode,ProdFloor,CreatedBy) values ({0},'{1}','{2}','{3}')", Convert.ToInt32(IppId), table1("OrgCode"), table1("ProdFloor"), LoginData.User.ToUpper)
                        da.ExecuteNonQuery(strHeader)
                    End If
                    strDetail = String.Format("INSERT INTO T_eJITIPPItem ([IPPID],[DJ],[Assembly],[IPPQty],[IPPDate],[Group],[Remarks])  values ({0},'{1}','{2}',{3},'{4}','{5}','{6}')", Convert.ToInt32(IppId), table1("DJ"), table1("Assembly"), table1("IPPQty"), table1("IPPDate"), table1("Group"), table1("Remarks"))
                    da.ExecuteNonQuery(strDetail)
                Next
                Upload_EJITIPP = True
            Catch ex As Exception
                ErrorLogging("UploadEJITIPP-SaveIPP", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Upload_EJITIPP = False
            End Try
        End Using
    End Function
    Public Function ValidEjitIpp(ByVal p_orgcode As String, ByVal p_dj As String, ByVal p_item As String, ByVal p_floor As String) As String
        ValidEjitIpp = ""
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim errmsg As String
            sda = da.Sda_Sele()
            Try
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then
                    sda.SelectCommand.Connection.Close()
                End If
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_valid_ejit_ipp"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@orgcode", SqlDbType.VarChar, 50).Value = p_orgcode
                sda.SelectCommand.Parameters.Add("@dj", SqlDbType.VarChar, 100).Value = p_dj
                sda.SelectCommand.Parameters.Add("@item", SqlDbType.VarChar, 150).Value = p_item
                sda.SelectCommand.Parameters.Add("@floor", SqlDbType.VarChar, 150).Value = p_floor
                sda.SelectCommand.Parameters.Add("@errormsg", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                ValidEjitIpp = sda.SelectCommand.Parameters("@errormsg").Value
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ValidEjitIpp = ex.Message & ex.Source
            End Try
        End Using
    End Function
    Public Function ChangeEjitIpp(ByVal p_floor As String) As String
        ChangeEjitIpp = ""
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim errmsg As String
            sda = da.Sda_Sele()
            Try
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then
                    sda.SelectCommand.Connection.Close()
                End If
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "ora_change_ipp"
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Parameters.Add("@prodfloor", SqlDbType.VarChar, 150).Value = p_floor
                sda.SelectCommand.Connection.Open()
                sda.SelectCommand.ExecuteNonQuery()
                sda.SelectCommand.Connection.Close()
                ChangeEjitIpp = "Y"
            Catch ex As Exception
                ChangeEjitIpp = "N"
            End Try
        End Using
    End Function
    Public Function checkIPPFormat(ByVal MRListData As DataSet, ByVal exceldata As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim myDataRow As Data.DataRow
        Dim Productionfloor As String = ""
        Dim r1, r2 As DataRow()
        Using da As DataAccess = GetDataAccess()
            Try
                Dim v_date1 As Date
                For Each dr As DataRow In exceldata.Tables(0).Rows
                    myDataRow = MRListData.Tables(0).NewRow()
                    myDataRow("OrgCode") = OracleERPLogin.OrgCode
                    myDataRow("ProdFloor") = IIf(dr(0) Is DBNull.Value, "", dr(0).ToString().Trim)
                    myDataRow("DJ") = IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim.ToUpper)
                    myDataRow("Assembly") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim.ToUpper)
                    myDataRow("IPPQty") = IIf(Microsoft.VisualBasic.IsDBNull(dr(3)), DBNull.Value, dr(3))
                    v_date1 = IIf(dr(4) Is DBNull.Value, "", FormatDateTime(dr(4), DateFormat.ShortDate))
                    myDataRow("Group") = IIf(dr(5) Is DBNull.Value, "", dr(5).ToString().Trim)
                    myDataRow("Remarks") = IIf(dr(6) Is DBNull.Value, "", dr(6).ToString().Trim)
                    If Microsoft.VisualBasic.IsDBNull(dr(3)) Then
                        myDataRow("ErrorMsg") = "IPPQTY can not be blank"
                    ElseIf dr(3) < 0 Then
                        myDataRow("ErrorMsg") = "IPPQTY can not less than 0"
                    Else
                        myDataRow("ErrorMsg") = ""
                    End If
                    If dr(5).ToString.Length = 0 Then
                        If myDataRow("ErrorMsg") = "" Or myDataRow("ErrorMsg") Is DBNull.Value Then
                            myDataRow("ErrorMsg") = "Group column can't be blank"
                        End If
                    Else
                        If InStr(dr(5), "G") = 0 Then
                            If myDataRow("ErrorMsg") = "" Or myDataRow("ErrorMsg") Is DBNull.Value Then
                                myDataRow("ErrorMsg") = "Group must begin with G"
                            End If
                        End If
                    End If
                    If Microsoft.VisualBasic.IsDate(dr(4)) = False Then
                        If myDataRow("ErrorMsg") = "" Or myDataRow("ErrorMsg") Is DBNull.Value Then
                            myDataRow("ErrorMsg") = "Data format is not correct"
                        End If
                    Else
                        myDataRow("IPPDate") = v_date1.ToShortDateString
                    End If
                    If dr(0).ToString.Length > 0 Then
                        Productionfloor = dr(0).ToString().Trim
                    End If
                    r1 = MRListData.Tables("errordata").Select("DJ='" & myDataRow("DJ") & "' and IPPDate='" & myDataRow("IPPDate") & "' and ProdFloor='" & myDataRow("ProdFloor") & "'")
                    If (r1.Length > 0) Then
                        If (r1.Length = 1) Then
                            r1(0)("ErrorMsg") = "duplicate data!"
                        End If
                        myDataRow("ErrorMsg") = "duplicate data!"
                        MRListData.Tables(0).Rows.Add(myDataRow)
                        Continue For
                    End If
                    MRListData.Tables(0).Rows.Add(myDataRow)
                Next
            Catch ex As Exception
                ErrorLogging("MMC-checkIPPFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
                myDataRow("ErrorMsg") = ex.Message & ex.Source
                MRListData.Tables(0).Rows.Add(myDataRow)
                Return MRListData
            End Try
            Dim ds As DataSet = New DataSet
            Dim ipprow() As DataRow = Nothing
            Dim ippsum() As DataRow = Nothing
            Dim Sqlstr As String
            Sqlstr = String.Format("Select OrgCode,DJ,IPPQty,IPPDate from v_ora_get_ippdata with (nolock) where ProdFloor = '" & Productionfloor & "'")
            ds = da.ExecuteDataSet(Sqlstr, "IPPHistory")
            Dim i As Integer
            Dim j As Integer
            Dim OrgCdoe As String
            Dim DJ As String
            Dim IPPDate As String
            Dim Adjust_ippdate As Date
            Dim IPPItemRow As DataRow
            Dim TotalQty As Double
            Dim DJQty As Double
            Dim Adjustrow() As DataRow = Nothing
            Dim AdjustData As DataTable
            Dim AdjustItemRow As DataRow
            Dim v_date2 As Date
            Try
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    v_date2 = ds.Tables(0).Rows(i)("IPPDate")
                    ds.Tables(0).Rows(i)("IPPDate") = v_date2.ToShortDateString
                Next
                ds.Tables(0).AcceptChanges()
                For i = 0 To MRListData.Tables(0).Rows.Count - 1
                    If MRListData.Tables(0).Rows(i)("errormsg") Is DBNull.Value Or MRListData.Tables(0).Rows(i)("errormsg") = "" Then
                        ipprow = Nothing
                        OrgCdoe = MRListData.Tables(0).Rows(i)("OrgCode")
                        DJ = MRListData.Tables(0).Rows(i)("DJ")
                        IPPDate = MRListData.Tables(0).Rows(i)("IPPDate")
                        ipprow = ds.Tables(0).Select(" OrgCode = '" & OrgCdoe & "' and DJ = '" & DJ & "' and IPPDate = '" & IPPDate & "'")
                        If ipprow.Length > 0 Then
                            For j = 0 To ipprow.Length - 1
                                ipprow(j)("IPPQty") = MRListData.Tables(0).Rows(i)("IPPQty")
                                ipprow(j).AcceptChanges()
                            Next
                            ds.Tables(0).AcceptChanges()
                        Else
                            IPPItemRow = ds.Tables(0).NewRow
                            IPPItemRow("OrgCode") = OrgCdoe
                            IPPItemRow("DJ") = DJ
                            IPPItemRow("IPPQty") = MRListData.Tables(0).Rows(i)("IPPQty")
                            IPPItemRow("IPPDate") = IPPDate
                            ds.Tables(0).Rows.Add(IPPItemRow)
                        End If
                        ds.Tables("IPPHistory").AcceptChanges()
                    End If
                Next
                For i = 0 To MRListData.Tables(0).Rows.Count - 1
                    If MRListData.Tables(0).Rows(i)("errormsg") Is DBNull.Value Or MRListData.Tables(0).Rows(i)("errormsg") = "" Then
                        MRListData.Tables(0).Rows(i)("errormsg") = ValidEjitIpp(OracleERPLogin.OrgCode, MRListData.Tables(0).Rows(i)("DJ"), MRListData.Tables(0).Rows(i)("Assembly"), MRListData.Tables(0).Rows(i)("ProdFloor"))
                    End If
                    OrgCdoe = MRListData.Tables(0).Rows(i)("OrgCode")
                    DJ = MRListData.Tables(0).Rows(i)("DJ")
                    ippsum = ds.Tables(0).Select(" OrgCode = '" & OrgCdoe & "' and DJ = '" & DJ & "'")
                    If ippsum.Length > 0 Then
                        TotalQty = 0
                        For j = 0 To ippsum.Length - 1
                            TotalQty = TotalQty + ippsum(j)("IPPQty")
                        Next
                        Sqlstr = String.Format("select POQty from T_DJInfo where OrgCode = '" & OrgCdoe & "' and PO = '" & DJ & "' and Model = PCBA ")
                        DJQty = Convert.ToDouble(da.ExecuteScalar(Sqlstr))
                        If TotalQty > DJQty Then
                            If MRListData.Tables(0).Rows(i)("errormsg") Is DBNull.Value Or MRListData.Tables(0).Rows(i)("errormsg") = "" Then
                                MRListData.Tables(0).Rows(i)("errormsg") = "Upload Excel IPPQty + History IPPQty greater than DJ " & DJ & " Start Qty " & DJQty
                            End If
                        End If
                    End If
                    IPPDate = MRListData.Tables(0).Rows(i)("IPPDate")
                    Sqlstr = String.Format("select dbo.ora_adjust_ippdate( '" & IPPDate & "')")
                    Adjust_ippdate = Convert.ToDateTime(da.ExecuteScalar(Sqlstr))
                    If IPPDate <> Adjust_ippdate Then
                        MRListData.Tables(0).Rows(i)("IPPDate") = Adjust_ippdate.ToShortDateString
                        MRListData.Tables(0).AcceptChanges()
                    End If
                Next
                AdjustData = MRListData.Tables(0).Clone
                AdjustData.TableName = MRListData.Tables(0).TableName
                For i = 0 To MRListData.Tables(0).Rows.Count - 1
                    OrgCdoe = MRListData.Tables(0).Rows(i)("OrgCode")
                    DJ = MRListData.Tables(0).Rows(i)("DJ")
                    IPPDate = MRListData.Tables(0).Rows(i)("IPPDate")
                    Adjustrow = AdjustData.Select(" OrgCode = '" & OrgCdoe & "' and DJ = '" & DJ & "' and IPPDate = '" & IPPDate & "'")
                    TotalQty = MRListData.Tables(0).Compute("Sum(IPPQty)", " OrgCode = '" & OrgCdoe & "' and DJ = '" & DJ & "' and IPPDate = '" & IPPDate & "'")
                    If Adjustrow.Length = 0 Then
                        AdjustItemRow = AdjustData.NewRow
                        AdjustItemRow("OrgCode") = OrgCdoe
                        AdjustItemRow("ProdFloor") = MRListData.Tables(0).Rows(i)("ProdFloor")
                        AdjustItemRow("DJ") = DJ
                        AdjustItemRow("Assembly") = MRListData.Tables(0).Rows(i)("Assembly")
                        AdjustItemRow("IPPQty") = TotalQty
                        AdjustItemRow("IPPDate") = IPPDate
                        AdjustItemRow("Group") = MRListData.Tables(0).Rows(i)("Group")
                        AdjustItemRow("Remarks") = MRListData.Tables(0).Rows(i)("Remarks")
                        AdjustItemRow("ErrorMsg") = MRListData.Tables(0).Rows(i)("ErrorMsg")
                        AdjustData.Rows.Add(AdjustItemRow)
                    End If
                Next
                MRListData.Tables.Clear()
                MRListData.Tables.Add(AdjustData)
                Return MRListData
            Catch oe As Exception
                ErrorLogging("UploadEJITIPP-CheckIPPFormat", "", oe.Message & oe.Source, "E")
                MRListData.Tables(0).Rows(i)("errormsg") = oe.Message & oe.Source
                Return MRListData
            End Try
        End Using
    End Function
#End Region

#Region "KanbanReplenishment"

    ''''//-----------------------------------------------------------------------

    ''''//  The Handhelp program "Kanban Replenishment" will invoke it

    ''''//  <copyright>Copyright (c) EMRSN. All rights reserved.</copyright>

    ''''//  <date>10-08-2010</date>

    ''''//  <author>Jackson Huang</author>

    ''''//  <comment>

    ''''// The Kanban data will load and upload to Oralce DB.

    ''''//  by hard code

    ''''//  </comment>

    ''''//-----------------------------------------------------------------------
    Public Function getKanbanIDinfo(ByVal ERPLoginData As ERPLogin, ByVal iKanbanNum As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim objDS As New DataSet()
            objDS.Tables.Add("KanbanItem")
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_kb_item_replenish"
                Dim oraclePara As OracleParameter = New OracleParameter("o_message", OracleType.VarChar, 240)


                oraclePara.Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_kb_item", OracleType.Cursor))
                ''oda.SelectCommand.Parameters.Add(New OracleParameter("o_message", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters.Add(oraclePara)

                oda.SelectCommand.Parameters("o_kb_item").Direction = ParameterDirection.Output
                ''oda.SelectCommand.Parameters("o_message").Direction = ParameterDirection.Output


                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_kanban_number", OracleType.VarChar, 240)).Value = iKanbanNum

                oda.SelectCommand.Connection.Open()
                oda.Fill(objDS, "KanbanItem")
                objDS.Tables(0).Columns.Add(New Data.DataColumn("o_message", System.Type.GetType("System.String")))

                objDS.Tables(0).Rows(0)("o_message") = oraclePara.Value

                oda.SelectCommand.Connection.Close()
                Return objDS
            Catch ex As Exception
                ErrorLogging("Get_KanbanData", ERPLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using
    End Function


    Public Function UpdateKanbanCodeDS(ByVal p_OracleLoginData As ERPLogin, ByVal iKanbanID As String, ByVal p_needbydate As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Dim iCnt As Integer
            Dim strMsg As String = ""
            Dim buildPlanId As Int64 = -1
            Dim strOrgCode = p_OracleLoginData.OrgCode
            Try
                If comm.Connection.State = ConnectionState.Closed Then
                    comm.Connection.Open()
                End If


                If (Int64.TryParse(Me.GetNextBPID(p_OracleLoginData), buildPlanId)) Then

                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.XXETR_KanBan_pkg.kb_replenish"
                    comm.Parameters.Add("p_user_id", OracleType.Number).Value = p_OracleLoginData.UserID '15891

                    ''''"Hard code the Kanban ResID and AppID"''''

                    ''If strOrgCode = "365" OrElse strOrgCode = "355" OrElse strOrgCode = "061" OrElse strOrgCode = "084" Then
                    ''    comm.Parameters.Add("p_kb_resp_id", OracleType.Number).Value = 55022
                    ''ElseIf strOrgCode = "019" OrElse strOrgCode = "063" OrElse strOrgCode = "086" OrElse strOrgCode = "404" OrElse strOrgCode = "405" OrElse strOrgCode = "406" Then
                    ''    comm.Parameters.Add("p_kb_resp_id", OracleType.Number).Value = 53790      'the 404 setting
                    ''ElseIf strOrgCode = "411" OrElse strOrgCode = "410" OrElse strOrgCode = "409" OrElse strOrgCode = "408" OrElse strOrgCode = "064" OrElse strOrgCode = "412" OrElse strOrgCode = "413" Then
                    ''    comm.Parameters.Add("p_kb_resp_id", OracleType.Number).Value = 54714
                    ''ElseIf strOrgCode = "416" OrElse strOrgCode = "065" OrElse strOrgCode = "417" OrElse strOrgCode = "415" Then
                    ''    comm.Parameters.Add("p_kb_resp_id", OracleType.Number).Value = 54733
                    ''End If
                    ''comm.Parameters.Add("p_kb_appl_id", OracleType.Number).Value = 401        'the 404 setting
                    ''''''''''''''''''''''''''''''''''''''''''''''


                    comm.Parameters.Add("p_kb_resp_id", OracleType.Number).Value = p_OracleLoginData.RespID_KB
                    comm.Parameters.Add("p_kb_appl_id", OracleType.Number).Value = p_OracleLoginData.AppID_KB
                    'comm.Parameters.Add("p_po_resp_id", OracleType.Number).Value = p_OracleLoginData.RespID_PO  '53524'53519 'p_OracleLoginData.RespID_PO
                    'comm.Parameters.Add("p_po_appl_id", OracleType.Number).Value = p_OracleLoginData.AppID_PO '201
                    comm.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = p_OracleLoginData.OrgCode '"404" 'p_OracleLoginData.OrgCode
                    comm.Parameters.Add("p_kanban_number", OracleType.VarChar, 240).Value = iKanbanID    'int
                    comm.Parameters.Add("p_need_by_date", OracleType.VarChar).Value = p_needbydate
                    ErrorLogging("Transfer_KanbanData", p_OracleLoginData.User, buildPlanId, "I")
                    comm.Parameters.Add("p_kb_identity", OracleType.Number, 20).Value = buildPlanId
                    ErrorLogging("Transfer_KanbanData", p_OracleLoginData.User, buildPlanId, "I")
                    comm.Parameters.Add("o_message", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                    iCnt = comm.ExecuteNonQuery()
                    strMsg = comm.Parameters("o_message").Value.ToString()
                    ErrorLogging("Transfer_KanbanData", p_OracleLoginData.User, buildPlanId, "I")

                    If strMsg = "Normal completion" OrElse strMsg = "" Then
                        comm.Transaction.Commit()
                        SkipLot(p_OracleLoginData, buildPlanId)
                    Else
                        comm.Transaction.Rollback()
                    End If

                End If


            Catch ex As Exception
                ErrorLogging("Transfer_KanbanData", p_OracleLoginData.User, ex.Message & ex.Source, "E")
                Throw ex
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
            End Try
            Return strMsg
        End Using

    End Function


    ''' <summary>
    ''' Invokes package to get data from Oralce.
    ''' Then Invokes SaveData function into eTrace datatable.
    ''' </summary>
    ''' <param name="oracleLoginData">ERPLogin</param>
    ''' <param name="buildPlanId">Int64</param>
    ''' <returns>if <c>True</c> the skiplot operation is successful(include don't need skiplot),
    ''' otherwish returns <c>false</c></returns>
    ''' <remarks></remarks>
    Public Function SkipLot(ByVal oracleLoginData As ERPLogin, ByVal buildPlanId As Int64) As Boolean

        Using da As DataAccess = GetDataAccess()
            Dim dsKanbanSkipLot As DataSet = New DataSet()
            Dim oracleDataAdapter As OracleDataAdapter = da.Oda_Sele()

            Dim strMsg As String = String.Empty
            ''Dim strOrgCode = oracleLoginData.OrgCode

            Try

                oracleDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oracleDataAdapter.SelectCommand.CommandText = "apps.xxetr_kanban_pkg.process_ejit_skiplot"

                ''            ''                    (
                ''            '' p_org_code         IN VARCHAR2,
                ''            '' p_user_id          IN NUMBER,
                ''            '' p_kb_identity      IN NUMBER,
                ''            '' o_buildplanitem    OUT cur_kb,
                ''            '' o_kbpickinglist    OUT cur_kb,
                ''            '' o_return_flag      OUT VARCHAR2,
                ''            '' o_return_message   OUT VARCHAR2
                ''            '');

                ''ErrorLogging("Transfer_KanbanData", oracleLoginData.User, oracleLoginData.OrgCode.ToString() + " " + oracleLoginData.UserID.ToString() + " " + buildPlanId.ToString() + " ", "I")
                oracleDataAdapter.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = oracleLoginData.OrgCode '"404" 'p_OracleLoginData.OrgCode
                oracleDataAdapter.SelectCommand.Parameters.Add("p_user_id", OracleType.Int32).Value = oracleLoginData.UserID '1589
                oracleDataAdapter.SelectCommand.Parameters.Add("p_kb_identity", OracleType.Number, 20).Value = buildPlanId
                oracleDataAdapter.SelectCommand.Parameters.Add("o_buildplanitem", OracleType.Cursor).Direction = ParameterDirection.Output
                oracleDataAdapter.SelectCommand.Parameters.Add("o_kbpickinglist", OracleType.Cursor).Direction = ParameterDirection.Output
                oracleDataAdapter.SelectCommand.Parameters.Add("o_return_flag", OracleType.VarChar, 200).Direction = ParameterDirection.Output
                oracleDataAdapter.SelectCommand.Parameters.Add("o_return_message", OracleType.VarChar, 200).Direction = ParameterDirection.Output

                oracleDataAdapter.SelectCommand.Connection.Open()
                oracleDataAdapter.Fill(dsKanbanSkipLot)
                oracleDataAdapter.SelectCommand.Connection.Close()


                strMsg = oracleDataAdapter.SelectCommand.Parameters("o_return_message").Value.ToString()

                If (strMsg = "Normal completion" OrElse strMsg = String.Empty) Then

                    If (Not IsNothing(dsKanbanSkipLot)) Then

                        If (2 = dsKanbanSkipLot.Tables.Count) Then
                            If (dsKanbanSkipLot.Tables(0).Rows.Count <> 0 And dsKanbanSkipLot.Tables(1).Rows.Count <> 0) Then

                                dsKanbanSkipLot.Tables(0).TableName = "BuildPlan"
                                dsKanbanSkipLot.Tables(1).TableName = "requireList_table"

                                If (SaveData(dsKanbanSkipLot, oracleLoginData)) Then
                                    Return True
                                End If
                            End If
                        End If

                    End If
                End If
            Catch ex As Exception
                ErrorLogging("KanbanSkipLot", oracleLoginData.User, ex.Message & ex.Source, "E")
                Return False
            Finally
                If oracleDataAdapter.SelectCommand.Connection.State <> ConnectionState.Closed Then oracleDataAdapter.SelectCommand.Connection.Close()
            End Try
            Return True
        End Using

    End Function


#End Region
#Region "BuildPlan"
    Public Function CheckFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim Productionfloor As String = ""
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Try
                For Each dr As DataRow In ExcelData.Tables(0).Rows
                    RowIndex = RowIndex + 1
                    If (FixNull(dr(0)) = "" And FixNull(dr(1)) = "" And FixNull(dr(2)) = "" And FixNull(dr(3)) = "" And FixNull(dr(4)) = "" And FixNull(dr(5)) = "" And FixNull(dr(6)) = "") Then
                        Continue For
                    End If

                    myDataRow = MRListData.Tables("BuildPlan").NewRow()
                    myDataRow("Line") = RowIndex
                    myDataRow("BuildPlanIdentity") = InvID
                    myDataRow("OrgCode") = OracleERPLogin.OrgCode

                    If FixNull(dr(0)) = "" Then
                        myDataRow("ErrorMsg") = "Error: DJNo must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(1)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Model must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(2)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Subinventory must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(3)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Quantity must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(4)) = "" Then
                        myDataRow("ErrorMsg") = "Error: DeliveryDate must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(5)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Productionfloor must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(6)) = "" Then
                        myDataRow("ErrorMsg") = "Error: WipQty must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If Double.TryParse(dr(3), d) = False OrElse dr(3) < 0 Then
                        myDataRow("ErrorMsg") = "Error: Enter Quantity only with Positive number"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If Double.TryParse(dr(6), d) = False OrElse dr(6) < 0 Then
                        myDataRow("ErrorMsg") = "Error: Enter WipQTY only with Positive number"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    dr(0) = dr(0).ToString.Trim.ToUpper
                    myDataRow("DJNo") = IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim)
                    myDataRow("model") = IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim)
                    myDataRow("Subinventory") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                    myDataRow("Productionfloor") = IIf(dr(5) Is DBNull.Value, "", dr(5).ToString.Trim)
                    myDataRow("Quantity") = IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim)
                    myDataRow("DeliveryDate") = IIf(dr(4) Is DBNull.Value, "", dr(4).ToString().Trim)
                    myDataRow("WipQty") = IIf(dr(6) Is DBNull.Value, "", dr(6).ToString.Trim)

                    If (myDataRow("Productionfloor").ToString.Contains("…")) Then
                        myDataRow("Productionfloor") = myDataRow("Productionfloor").ToString.Replace("…", "...")
                    End If

                    myDataRow("DeliveryDate") = dr(4).ToString.Split(" ")(0)


                    If (FixNull(dr(4)) <> "" AndAlso Date.TryParse(dr(4), dt) = False) Or myDataRow("DeliveryDate").ToString.Substring(myDataRow("DeliveryDate").ToString.LastIndexOf("/") + 1).Length <> 4 Then
                        myDataRow("ErrorMsg") = "Error: Enter DeliveryDate with format: MM/DD/YYYY"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For


                    End If

                    If (Productionfloor = "") Then
                        Productionfloor = myDataRow("Productionfloor")
                    ElseIf (Productionfloor <> myDataRow("Productionfloor")) Then
                        myDataRow("ErrorMsg") = "Error: Productionfloor must same as other line"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If


                    'Dim _eTrace_TDCFunction As New TestDataCollection()

                    ' Status =1 ,release DJ
                    'If (Not _eTrace_TDCFunction.isValidDJ(dr(0).ToString.Trim, "3", OracleERPLogin)) Then
                    '    myDataRow("ErrorMsg") = "Error: DJNo is not valid!"
                    '    MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                    '    Continue For
                    'End If

                    Dim msg As String = isValidDJModel(dr(0).ToString.Trim, dr(1).ToString.Trim, dr(3), dr(6), OracleERPLogin)
                    If (msg.Length > 0) Then
                        myDataRow("ErrorMsg") = msg
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If (Not isValidLocSubInv(dr(2).ToString.Trim, myDataRow("Productionfloor"), OracleERPLogin)) Then
                        myDataRow("ErrorMsg") = "Error: SubInventory and productionfloor are not match!"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    Dim datenow As Date
                    Dim DifOfDate As Integer

                    datenow = Now.Date
                    DifOfDate = DateDiff("d", datenow, myDataRow("DeliveryDate"))
                    If DifOfDate < 0 Then
                        myDataRow("ErrorMsg") = "Error: Exp Date " & myDataRow("DeliveryDate") & " is not in correct interval"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                Next

                Return MRListData

            Catch ex As Exception
                ErrorLogging("KanbanClient-AddBuildPlan", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function


    Public Function CheckPNFormat(ByVal InvID As String, ByVal MRListData As DataSet, ByVal ExcelData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet
        Dim RowIndex As Integer = 0
        Dim myDataRow As Data.DataRow
        Dim d As Double
        Dim dt As Date
        Dim Productionfloor As String = ""
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Try
                For Each dr As DataRow In ExcelData.Tables(0).Rows
                    If (FixNull(dr(0)) = "" And FixNull(dr(1)) = "" And FixNull(dr(2)) = "" And FixNull(dr(3)) = "" And FixNull(dr(4)) = "") Then
                        Continue For
                    End If
                    RowIndex = RowIndex + 1
                    myDataRow = MRListData.Tables("BuildPlan").NewRow()
                    myDataRow("Line") = RowIndex
                    myDataRow("BuildPlanIdentity") = InvID
                    myDataRow("OrgCode") = OracleERPLogin.OrgCode

                    myDataRow("PN") = IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim)
                    myDataRow("Subinventory") = IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim)
                    myDataRow("Quantity") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                    myDataRow("DeliveryDate") = IIf(dr(3) Is DBNull.Value, "", dr(3).ToString().Trim)
                    myDataRow("Productionfloor") = IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim)

                    myDataRow("DJNO") = "Manual Upload"
                    myDataRow("Model") = myDataRow("PN")
                    myDataRow("DeliveryDate") = dr(3).ToString.Split(" ")(0)
                    myDataRow("Shippinginstruction") = dr(5).ToString.Trim

                    If (myDataRow("Productionfloor").ToString.Contains("…")) Then
                        myDataRow("Productionfloor") = myDataRow("Productionfloor").ToString.Replace("…", "...")
                    End If
                    If FixNull(dr(0)) = "" Then
                        myDataRow("ErrorMsg") = "Error: KanBanItem must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(1)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Subinventory must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(2)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Quantity must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(3)) = "" Then
                        myDataRow("ErrorMsg") = "Error: DeliveryDate must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    ElseIf FixNull(dr(4)) = "" Then
                        myDataRow("ErrorMsg") = "Error: Productionfloor must be entered"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If Double.TryParse(dr(2), d) = False OrElse dr(2) <= 0 Then
                        myDataRow("ErrorMsg") = "Error: Enter Quantity only with Positive number"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If (FixNull(dr(3)) <> "" AndAlso Date.TryParse(dr(3), dt) = False) Or myDataRow("DeliveryDate").ToString.Substring(myDataRow("DeliveryDate").ToString.LastIndexOf("/") + 1).Length <> 4 Then
                        myDataRow("ErrorMsg") = "Error: Enter DeliveryDate with format: MM/DD/YYYY"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If




                    If (Productionfloor = "") Then
                        Productionfloor = myDataRow("Productionfloor")
                    ElseIf (Productionfloor <> myDataRow("Productionfloor")) Then
                        myDataRow("ErrorMsg") = "Error: Productionfloor must same as other line"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If


                    Dim _eTrace_TDCFunction As New TestDataCollection()


                    If (Not isValidLocSubInv(dr(1).ToString.Trim, myDataRow("Productionfloor"), OracleERPLogin)) Then
                        myDataRow("ErrorMsg") = "Error: SubInventory and productionfloor are not match!"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    Dim datenow As Date
                    Dim DifOfDate As Integer

                    datenow = Now.Date
                    DifOfDate = DateDiff("d", datenow, myDataRow("DeliveryDate"))
                    If DifOfDate < 0 Then
                        myDataRow("ErrorMsg") = "Error: Exp Date " & myDataRow("DeliveryDate") & " is not in correct interval"
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    If (Not isValideJITItem(myDataRow("PN"), OracleERPLogin)) Then
                        myDataRow("ErrorMsg") = "not valid ejit Item :" & myDataRow("PN").ToString
                        MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                        Continue For
                    End If

                    MRListData.Tables("BuildPlan").Rows.Add(myDataRow)
                Next

                Return MRListData

            Catch ex As Exception
                ErrorLogging("KanbanClient-CheckKanbanItem", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Private Function isValideJITItem(ByVal itemname As String, ByVal LoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            isValideJITItem = False
            Dim ds As DataSet = New DataSet
            Dim oda As OracleDataAdapter = da.Oda_Sele
            Dim comm_submit As New OracleCommand()

            Try


                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.xxetr_po_detail_pkg.validate_ejit_item"
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_flag", OracleType.VarChar, 10))
                oda.SelectCommand.Parameters("o_flag").Direction = ParameterDirection.InputOutput

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_num", OracleType.VarChar, 240)).Value = itemname
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = LoginData.OrgCode

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "isValideJITItem")
                oda.SelectCommand.Connection.Close()
                Dim Flag As String

                Flag = oda.SelectCommand.Parameters("o_flag").Value.ToString
                If (Flag = "Y") Then
                    isValideJITItem = True
                End If

            Catch oe As Exception
                ErrorLogging("Kanban-isValidDJModel", LoginData.User, oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function CommitBuildPlan(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin, ByVal uploadType As Integer, ByVal saType As Integer, ByVal triggerType As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter
            Dim sda As SqlClient.SqlDataAdapter
            Dim buildplan_type As String
            Dim Sqlstr, KBFlag As String
            Try
                oda = da.Oda_Insert()
                buildplan_type = "oda"
                oda.InsertCommand.CommandType = CommandType.StoredProcedure
                If (uploadType = 1) Then
                    oda.InsertCommand.CommandText = "APPS.XXETR_KanBan_pkg.ins_kb_dj_v2"
                    oda.InsertCommand.Parameters.Add("p_dj", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_model", OracleType.VarChar, 50)
                    oda.InsertCommand.Parameters.Add("p_trigger_type", OracleType.VarChar, 50)
                    oda.InsertCommand.Parameters.Add("o_message", OracleType.VarChar, 1000)
                ElseIf (uploadType = 2) Then
                    oda.InsertCommand.CommandText = "APPS.XXETR_KanBan_pkg.ins_kb_pn"
                    oda.InsertCommand.Parameters.Add("p_pn", OracleType.VarChar, 25)
                    oda.InsertCommand.Parameters.Add("p_ship_instruction", OracleType.VarChar, 480)
                End If
                oda.InsertCommand.Parameters.Add("p_kb_identity", OracleType.VarChar, 25)
                oda.InsertCommand.Parameters.Add("p_org_code", OracleType.VarChar, 25)
                oda.InsertCommand.Parameters.Add("p_subinventory", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_production_line", OracleType.VarChar)
                oda.InsertCommand.Parameters.Add("p_qty", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_delivery_date", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_wip_qty", OracleType.Int32)
                'oda.InsertCommand.Parameters.Add("o_message", OracleType.VarChar, 240)

                'oda.InsertCommand.Parameters("o__message").Direction = ParameterDirection.InputOutput
                If (uploadType = 1) Then
                    oda.InsertCommand.Parameters("p_dj").SourceColumn = "DJNO"
                    oda.InsertCommand.Parameters("p_model").SourceColumn = "Model"
                    oda.InsertCommand.Parameters("p_trigger_type").Value = triggerType
                    oda.InsertCommand.Parameters("o_message").SourceColumn = "ErrorMsg"
                    oda.InsertCommand.Parameters("o_message").Direction = ParameterDirection.Output
                ElseIf (uploadType = 2) Then
                    oda.InsertCommand.Parameters("p_PN").SourceColumn = "PN"
                    oda.InsertCommand.Parameters("p_ship_instruction").SourceColumn = "Shippinginstruction"
                End If
                oda.InsertCommand.Parameters("p_kb_identity").SourceColumn = "BuildPlanIdentity"
                oda.InsertCommand.Parameters("p_org_code").SourceColumn = "Orgcode"
                oda.InsertCommand.Parameters("p_subinventory").SourceColumn = "Subinventory"
                oda.InsertCommand.Parameters("p_production_line").SourceColumn = "Productionfloor"
                oda.InsertCommand.Parameters("p_qty").SourceColumn = "Quantity"
                oda.InsertCommand.Parameters("p_delivery_date").SourceColumn = "DeliveryDate"
                oda.InsertCommand.Parameters("p_wip_qty").SourceColumn = "WipQty"
                'oda.InsertCommand.Parameters("o_message").SourceColumn = "o_message"



                oda.InsertCommand.Connection.Open()
                oda.Update(p_ds.Tables("buildplan"))
                oda.InsertCommand.Connection.Close()
                'oda.Dispose()

                If Not p_ds Is Nothing AndAlso p_ds.Tables.Count > 0 AndAlso p_ds.Tables("BuildPlan").Rows.Count > 0 Then
                    Dim DR() As DataRow = Nothing
                    DR = p_ds.Tables("BuildPlan").Select("ErrorMsg <> ''")
                    If DR.Length <> 0 Then
                        Return p_ds
                    End If
                End If

            Catch oe As Exception
                ErrorLogging("CommitBuildPlan", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try

            Try
                Dim i As Integer
                For i = 0 To p_ds.Tables(0).Rows.Count - 1
                    If p_ds.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        p_ds.Tables(0).Rows(i).SetAdded()
                    End If
                Next

                'oda = da.Oda_Sele()
                'oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'If (uploadType = 1) Then
                '    oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_kb_item_by_wip"
                '    oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_type", OracleType.Int16))
                '    oda.SelectCommand.Parameters("p_item_type").Value = saType
                '    oda.SelectCommand.Parameters.Add(New OracleParameter("p_trigger_type", OracleType.VarChar, 25))
                '    oda.SelectCommand.Parameters("p_trigger_type").Value = triggerType
                'ElseIf (uploadType = 2) Then
                '    oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_kb_item_by_PN"
                'End If
                'oda.SelectCommand.Parameters.Add(New OracleParameter("p_kb_identity", OracleType.VarChar, 25))
                'oda.SelectCommand.Parameters("p_kb_identity").Value = p_ds.Tables("BuildPlan").Rows(0)("BuildPlanIdentity")
                'oda.SelectCommand.Parameters.Add(New OracleParameter("o_kb_item", OracleType.Cursor))
                'oda.SelectCommand.Parameters.Add(New OracleParameter("o_message", OracleType.VarChar, 240))
                'oda.SelectCommand.Parameters("o_kb_item").Direction = ParameterDirection.Output
                'oda.SelectCommand.Parameters("o_message").Direction = ParameterDirection.Output

                'oda.SelectCommand.Connection.Open()
                'oda.Fill(p_ds, "requireList_table")
                'oda.SelectCommand.Connection.Close()
                Sqlstr = String.Format("Select Value from T_Config where ConfigID = 'KAB001'")
                KBFlag = Convert.ToString(da.ExecuteScalar(Sqlstr))

                'Valiadte upload buildplan data by DJ logic (0=Normal Call Oracle package, 1=Call sql server store procedure)

                If (uploadType = 1) Then
                    If KBFlag = "0" Then
                        oda = da.Oda_Sele()
                        buildplan_type = "oda"
                        oda.SelectCommand.CommandType = CommandType.StoredProcedure
                        oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_kb_item_by_wip"
                        oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_type", OracleType.Int16))
                        oda.SelectCommand.Parameters("p_item_type").Value = saType
                        oda.SelectCommand.Parameters.Add(New OracleParameter("p_trigger_type", OracleType.VarChar, 25))
                        oda.SelectCommand.Parameters("p_trigger_type").Value = triggerType

                        oda.SelectCommand.Parameters.Add(New OracleParameter("p_kb_identity", OracleType.VarChar, 25))
                        oda.SelectCommand.Parameters("p_kb_identity").Value = p_ds.Tables("BuildPlan").Rows(0)("BuildPlanIdentity")
                        oda.SelectCommand.Parameters.Add(New OracleParameter("o_kb_item", OracleType.Cursor))
                        oda.SelectCommand.Parameters.Add(New OracleParameter("o_message", OracleType.VarChar, 240))
                        oda.SelectCommand.Parameters("o_kb_item").Direction = ParameterDirection.Output
                        oda.SelectCommand.Parameters("o_message").Direction = ParameterDirection.Output

                        oda.SelectCommand.Connection.Open()
                        oda.Fill(p_ds, "requireList_table")
                        oda.SelectCommand.Connection.Close()
                    Else
                        Dim ds As New DataSet
                        sda = da.Sda_Sele()
                        buildplan_type = "sda"
                        sda.SelectCommand.CommandType = CommandType.StoredProcedure
                        sda.SelectCommand.CommandText = "ora_get_kb_item_by_wip"
                        sda.SelectCommand.CommandTimeout = TimeOut_M5

                        sda.SelectCommand.Parameters.Add("@p_kb_identity", SqlDbType.BigInt).Value = p_ds.Tables("BuildPlan").Rows(0)("BuildPlanIdentity")
                        sda.SelectCommand.Parameters.Add("@p_item_type", SqlDbType.Int).Value = saType
                        sda.SelectCommand.Parameters.Add("@p_trigger_type", SqlDbType.VarChar, 150).Value = triggerType

                        sda.SelectCommand.Connection.Open()
                        sda.Fill(ds)
                        sda.SelectCommand.Connection.Close()

                        ds.Tables(0).TableName = "skip_table"
                        ds.Tables(1).TableName = "requireList_table"


                        p_ds.Tables("requireList_table").Clear()
                        p_ds.Tables("requireList_table").Merge(ds.Tables("requireList_table"))
                        For i = 0 To ds.Tables(0).Rows.Count - 1
                            Dim oda_2 As OracleCommand = da.OraCommand()
                            oda_2.CommandType = CommandType.StoredProcedure
                            oda_2.CommandText = "apps.xxetr_oe_pkg.exec_sql"

                            Sqlstr = " insert into xxetr.xxetr_ejit_skiplot (kb_identity,skip_lot_flag,skip_source_type,ejit_id,supplier_id,item_no,inventory_item_id," _
                               & "description,primary_unit_of_measure,item_type,quantity_per_assembly,kb_qty,required_quantity,quantity_issued,mpq,on_hand_qty," _
                               & "available_qty,mc_available_qty,wip_qty,open_irpo_qty,shortage,delivery_date,subinventory,source_organization_code," _
                               & "source_subinventory,error_message,ejit_trigger_type) values ( " _
                               & p_ds.Tables("BuildPlan").Rows(0)("BuildPlanIdentity") & "," _
                               & "'" & ds.Tables("skip_table").Rows(i)("skip_lot_flag") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("skip_source_type") & "'," _
                               & ds.Tables("skip_table").Rows(i)("ejit_id") & "," _
                               & "'" & ds.Tables("skip_table").Rows(i)("supplier_id") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("item_no") & "'," _
                               & ds.Tables("skip_table").Rows(i)("inventory_item_id") & "," _
                               & "'" & ds.Tables("skip_table").Rows(i)("description") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("primary_unit_of_measure") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("item_type") & "'," _
                               & ds.Tables("skip_table").Rows(i)("quantity_per_assembly") & "," _
                               & ds.Tables("skip_table").Rows(i)("kb_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("required_quantity") & "," _
                               & ds.Tables("skip_table").Rows(i)("quantity_issued") & "," _
                               & ds.Tables("skip_table").Rows(i)("mpq") & "," _
                               & ds.Tables("skip_table").Rows(i)("on_hand_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("available_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("mc_available_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("wip_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("open_irpo_qty") & "," _
                               & ds.Tables("skip_table").Rows(i)("shortage") & "," _
                               & "'" & ds.Tables("skip_table").Rows(i)("delivery_date") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("subinventory") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("source_organization_code") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("source_subinventory") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("error_message") & "'," _
                               & "'" & ds.Tables("skip_table").Rows(i)("ejit_trigger_type") & "' ) "
                            oda_2.Parameters.Add("p_strsql", OracleType.VarChar, 3000).Value = Sqlstr
                            oda_2.Connection.Open()
                            oda_2.ExecuteNonQuery()
                            oda_2.Connection.Close()
                        Next
                    End If

                ElseIf (uploadType = 2) Then
                    oda = da.Oda_Sele()
                    buildplan_type = "oda"
                    oda.SelectCommand.CommandType = CommandType.StoredProcedure
                    oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_kb_item_by_PN"
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_kb_identity", OracleType.VarChar, 25))
                    oda.SelectCommand.Parameters("p_kb_identity").Value = p_ds.Tables("BuildPlan").Rows(0)("BuildPlanIdentity")
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_kb_item", OracleType.Cursor))
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_message", OracleType.VarChar, 240))
                    oda.SelectCommand.Parameters("o_kb_item").Direction = ParameterDirection.Output
                    oda.SelectCommand.Parameters("o_message").Direction = ParameterDirection.Output

                    oda.SelectCommand.Connection.Open()
                    oda.Fill(p_ds, "requireList_table")
                    oda.SelectCommand.Connection.Close()
                End If

                Dim shortage, available As Decimal
                Dim myDataRow As Data.DataRow
                For i = 0 To p_ds.Tables("requireList_table").Rows.Count - 1

                    If (uploadType = 1) Then
                        available = p_ds.Tables("requireList_table").Rows(i)("ON_HAND_QTY") + p_ds.Tables("requireList_table").Rows(i)("mc_available_qty") + p_ds.Tables("requireList_table").Rows(i)("mc_allocated_qty") + p_ds.Tables("requireList_table").Rows(i)("open_irpo_qty") - p_ds.Tables("requireList_table").Rows(i)("wip_Qty")
                    ElseIf (uploadType = 2) Then
                        available = 0
                        p_ds.Tables("requireList_table").Rows(i)("mc_allocated_qty") = 0
                    End If

                    If (available > 0) Then
                        p_ds.Tables("requireList_table").Rows(i)("AVAILABLE_QTY") = available
                    Else
                        p_ds.Tables("requireList_table").Rows(i)("AVAILABLE_QTY") = 0
                    End If




                    If (p_ds.Tables("requireList_table").Rows(i)("MPQ") Is DBNull.Value) OrElse (p_ds.Tables("requireList_table").Rows(i)("MPQ") = 0) Then
                        p_ds.Tables("requireList_table").Rows(i)("MPQ") = 1
                    End If

                    shortage = Math.Ceiling((p_ds.Tables("requireList_table").Rows(i)("KB_qty") - p_ds.Tables("requireList_table").Rows(i)("AVAILABLE_QTY")) / p_ds.Tables("requireList_table").Rows(i)("MPQ")) * p_ds.Tables("requireList_table").Rows(i)("MPQ")

                    If (shortage < 0) Then
                        p_ds.Tables("requireList_table").Rows(i)("shortage") = 0
                    Else
                        p_ds.Tables("requireList_table").Rows(i)("shortage") = shortage
                    End If


                    myDataRow = p_ds.Tables("bomdata").NewRow()
                    myDataRow("Kanbam Item") = p_ds.Tables("requireList_table").Rows(i)("ITEM_NO")
                    myDataRow("Item Type") = p_ds.Tables("requireList_table").Rows(i)("ITEM_TYPE")
                    myDataRow("Source org code") = p_ds.Tables("requireList_table").Rows(i)("SOURCE_ORGANIZATION_CODE")
                    myDataRow("source subinventory code") = p_ds.Tables("requireList_table").Rows(i)("SOURCE_SUBINVENTORY")
                    myDataRow("MPQ") = p_ds.Tables("requireList_table").Rows(i)("MPQ")
                    myDataRow("UOM") = p_ds.Tables("requireList_table").Rows(i)("PRIMARY_UNIT_OF_MEASURE")
                    myDataRow("Own Floor Onhand") = p_ds.Tables("requireList_table").Rows(i)("ON_HAND_QTY")
                    myDataRow("Required Qty") = p_ds.Tables("requireList_table").Rows(i)("KB_QTY")
                    myDataRow("Available Qty") = p_ds.Tables("requireList_table").Rows(i)("AVAILABLE_QTY")
                    myDataRow("Order Qty") = p_ds.Tables("requireList_table").Rows(i)("SHORTAGE")
                    myDataRow("Delivery date") = p_ds.Tables("requireList_table").Rows(i)("DELIVERY_DATE")
                    If (p_ds.Tables("requireList_table").Rows(i)("ERROR_MESSAGE").ToString = "" And Decimal.Parse(p_ds.Tables("requireList_table").Rows(i)("mc_available_qty").ToString) > 0 And Decimal.Parse(p_ds.Tables("requireList_table").Rows(i)("SHORTAGE").ToString) > 0) Then
                        p_ds.Tables("requireList_table").Rows(i)("ERROR_MESSAGE") = "E:MC has on-hand, please create MO first! "
                    End If
                    myDataRow("MC Available Qty") = p_ds.Tables("requireList_table").Rows(i)("mc_available_qty")
                    myDataRow("MO Allocated Qty") = p_ds.Tables("requireList_table").Rows(i)("mc_allocated_qty")
                    myDataRow("Wip Qty") = p_ds.Tables("requireList_table").Rows(i)("wip_Qty")
                    myDataRow("open irpo qty") = p_ds.Tables("requireList_table").Rows(i)("open_irpo_qty")
                    myDataRow("Error message") = p_ds.Tables("requireList_table").Rows(i)("ERROR_MESSAGE")
                    myDataRow("ShippingInstruction") = p_ds.Tables("requireList_table").Rows(i)("shipping_instruction")
                    p_ds.Tables("bomdata").Rows.Add(myDataRow)
                Next



                Return p_ds

            Catch oex As OracleException
                ErrorLogging("ProcessCode:" & KBFlag & " ,Get_KanbanData", LoginData.User, oex.Message & oex.Source, "E")
                Throw oex
            Finally
                If buildplan_type = "sda" Then
                    If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
                Else
                    If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
                End If
            End Try



        End Using

    End Function

    Public Function SaveKanbanData(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim oda_h As OracleDataAdapter = da.Oda_Update_Tran()
            Dim errmsg As String = ""
            Dim aa As OracleString
            Dim i As Integer
            SaveKanbanData = ""

            Try
                'XXETR_SUPPLIER_QUOTA_SHARING
                oda_h.UpdateCommand.CommandType = CommandType.StoredProcedure
                oda_h.UpdateCommand.CommandText = "APPS.XXETR_KanBan_pkg.populate_pr_inte"
                oda_h.UpdateCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_user_id", OracleType.Int32)
                oda_h.UpdateCommand.Parameters.Add("p_item_id", OracleType.Int32)
                oda_h.UpdateCommand.Parameters.Add("p_qty", OracleType.Float)
                oda_h.UpdateCommand.Parameters.Add("p_delivery_date", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_subinventory", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_note_to_receiver", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_source_org_code", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_source_subinventory_code", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_item_no", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_item_type", OracleType.VarChar, 240)
                oda_h.UpdateCommand.Parameters.Add("p_ejit_id", OracleType.Number)
                oda_h.UpdateCommand.Parameters.Add("p_ship_instruction", OracleType.VarChar, 480)
                oda_h.UpdateCommand.Parameters.Add("o_message", OracleType.VarChar, 255)


                'p_ds.Tables(1).Columns.Add(New Data.DataColumn("org_code", System.Type.GetType("System.String")))
                'p_ds.Tables(1).Columns.Add(New Data.DataColumn("user_id", System.Type.GetType("System.String")))
                'p_ds.Tables(1).Columns.Add(New Data.DataColumn("message", System.Type.GetType("System.String")))
                'p_ds.Tables(1).Columns.Add(New Data.DataColumn("ProductionFloor", System.Type.GetType("System.String")))
                'For j As Integer = 0 To p_ds.Tables(1).Rows.Count - 1
                '    p_ds.Tables(1).Rows(j)("org_code") = LoginData.OrgCode
                '    p_ds.Tables(1).Rows(j)("user_id") = LoginData.UserID
                '    p_ds.Tables(1).Rows(j)("ProductionFloor") = p_ds.Tables(0).Rows(0)("ProductionFloor").ToString
                '    p_ds.Tables(1).Rows(j)("message") = ""

                'Next
                'Dim dr As DataRow() = p_ds.Tables(1).Select("shortage=0")

                'Dim size As Integer = dr.Length
                'For i = 0 To size - 1
                '    p_ds.Tables(1).Rows.Remove(dr(i))
                'Next


                oda_h.UpdateCommand.Parameters("o_message").Direction = ParameterDirection.InputOutput
                Dim dr As DataRow() = p_ds.Tables("requireList_table").Select("shortage<>0")

                For j As Integer = 0 To dr.Length - 1
                    oda_h.UpdateCommand.Parameters("p_org_code").Value = LoginData.OrgCode
                    oda_h.UpdateCommand.Parameters("p_user_id").Value = LoginData.UserID
                    oda_h.UpdateCommand.Parameters("p_item_id").Value = dr(j)("Inventory_item_id").ToString
                    oda_h.UpdateCommand.Parameters("p_qty").Value = dr(j)("Shortage").ToString
                    oda_h.UpdateCommand.Parameters("p_delivery_date").Value = dr(j)("Delivery_Date").ToString
                    oda_h.UpdateCommand.Parameters("p_subinventory").Value = dr(j)("subinventory").ToString
                    oda_h.UpdateCommand.Parameters("p_note_to_receiver").Value = p_ds.Tables("BuildPlan").Rows(0)("ProductionFloor").ToString
                    oda_h.UpdateCommand.Parameters("p_source_org_code").Value = dr(j)("source_organization_code").ToString
                    oda_h.UpdateCommand.Parameters("p_source_subinventory_code").Value = dr(j)("source_subinventory").ToString
                    oda_h.UpdateCommand.Parameters("p_item_no").Value = dr(j)("item_no").ToString
                    oda_h.UpdateCommand.Parameters("p_ejit_id").Value = dr(j)("ejit_id").ToString
                    oda_h.UpdateCommand.Parameters("p_item_type").Value = dr(j)("item_type").ToString
                    oda_h.UpdateCommand.Parameters("p_ship_instruction").Value = dr(j)("shipping_instruction").ToString
                    ' oda_h.UpdateCommand.Parameters("o_message").Value = dr(j)("message").ToString
                    oda_h.UpdateCommand.ExecuteNonQuery()
                    errmsg = oda_h.UpdateCommand.Parameters("o_message").Value.ToString
                    If (errmsg.Length <> 0) Then
                        ErrorLogging("SaveKanbanData-errmsg", LoginData.User.ToUpper, errmsg, "I")
                        oda_h.UpdateCommand.Transaction.Rollback()
                        oda_h.UpdateCommand.Connection.Close()
                        'oda_h.UpdateCommand.Connection.Dispose()
                        'oda_h.UpdateCommand.Dispose()
                        SaveKanbanData = errmsg
                        Exit Function
                    End If
                Next
                ' oda_h.UpdateCommand = comm
                'oda_h.Update(p_ds.Tables(1))

                '   Dim ndr As DataRow() = p_ds.Tables(1).Select("message <>''")

                '   If ndr.Length = 0 Then
                oda_h.UpdateCommand.Transaction.Commit()
                oda_h.UpdateCommand.Connection.Close()
                'oda_h.UpdateCommand.Connection.Dispose()
                'oda_h.UpdateCommand.Dispose()
                SaveData(p_ds, LoginData)
                '  Else
                'For i = 0 To (ndr.Length - 1)
                '    errmsg = ndr(i)("message").ToString
                '    ErrorLogging("SaveKanbanData-errmsg", LoginData.User.ToUpper, errmsg, "I")
                'Next

                'oda_h.UpdateCommand.Transaction.Rollback()
                'oda_h.UpdateCommand.Connection.Close()
                'oda_h.UpdateCommand.Connection.Dispose()
                'oda_h.UpdateCommand.Dispose()
                'SaveKanbanData = errmsg
                'Exit Function
                'End If

                'Threading.Thread.Sleep(5000)
            Catch oe As Exception
                SaveKanbanData = oe.Message
                ErrorLogging("SaveKanbanData", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                If (Not oda_h.UpdateCommand.Transaction Is Nothing) Then
                    oda_h.UpdateCommand.Transaction.Rollback()
                End If
            Finally
                If oda_h.UpdateCommand.Connection.State <> ConnectionState.Closed Then oda_h.UpdateCommand.Connection.Close()
                'oda_h.UpdateCommand.Connection.Dispose()
                'oda_h.UpdateCommand.Dispose()
            End Try
        End Using
    End Function


    Public Function SaveData(ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As Boolean
        SaveData = False
        Using da As DataAccess = GetDataAccess()

            Dim ra As Integer
            Dim strHeader, strDetail, strkanban, strsql As String
            Dim i As Integer
            Dim result1 As New Object
            Dim KBQty As Decimal = 0
            Dim MPQ As Decimal = 0
            Dim OnhandQty As Decimal = 0
            Dim mcqty As Decimal = 0
            Dim WipQty As Decimal = 0
            Dim Shippinginstruction As String

            Try
                'myConn.Open()
                strHeader = String.Format("INSERT INTO T_BuildPlanHeader (CreateBy) values ('{0}')", LoginData.User)
                da.ExecuteNonQuery(strHeader)

                strsql = String.Format("select max(BuildPlanId) from T_BuildPlanHeader", LoginData.User)
                Dim BuildPlanId As String = Convert.ToString(da.ExecuteScalar(strsql))

                'If (Not p_ds.Tables("BuildPlan").Columns.Contains("WipQty")) Then
                '    WipQty = -1
                'End If
                'If (Not p_ds.Tables("BuildPlan").Columns.Contains("Shippinginstruction")) Then
                'Shippinginstruction = ""
                'End If

                For i = 0 To p_ds.Tables("BuildPlan").Rows.Count - 1
                    Dim table1 As DataRow = p_ds.Tables("BuildPlan").Rows(i)
                    If (Not p_ds.Tables("BuildPlan").Columns.Contains("WipQty")) Then
                        WipQty = 0
                    Else
                        If (table1("WipQty") IsNot DBNull.Value) Then
                            WipQty = table1("WipQty")
                        Else
                            WipQty = 0
                        End If
                    End If
                    If (Not p_ds.Tables("BuildPlan").Columns.Contains("Shippinginstruction")) Then
                        Shippinginstruction = ""
                    Else
                        If (table1("Shippinginstruction") IsNot DBNull.Value) Then
                            Shippinginstruction = table1("Shippinginstruction").ToString.Replace("'", "")
                        Else
                            Shippinginstruction = ""
                        End If
                    End If

                        strDetail = String.Format("INSERT INTO T_BuildPlanItem (BuildPlanId,DJNO,model,Subinventory,RequireQty,DeliveryDate,Productionfloor,OrgCode,WipQty,ShipInstruction)  values ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", BuildPlanId, table1("DJNo"), table1("model"), table1("Subinventory"), table1("Quantity"), table1("DeliveryDate"), table1("Productionfloor"), table1("OrgCode"), WipQty, Shippinginstruction)
                        da.ExecuteNonQuery(strDetail)
                Next

                For i = 0 To p_ds.Tables(1).Rows.Count - 1
                    Dim table2 As DataRow = p_ds.Tables("requireList_table").Rows(i)

                    If p_ds.Tables("requireList_table").Columns.IndexOf("mc_available_qty") = -1 Then
                        mcqty = 0
                    Else
                        mcqty = table2("mc_available_qty")
                    End If

                    If (table2("KB_Qty") IsNot DBNull.Value) Then
                        KBQty = table2("KB_Qty")
                    End If
                    If (table2("MPQ") IsNot DBNull.Value) Then
                        MPQ = table2("MPQ")
                    End If
                    If (table2("On_hand_Qty") IsNot DBNull.Value) Then
                        OnhandQty = table2("On_hand_Qty")
                    End If
                    If (Not p_ds.Tables("requireList_table").Columns.Contains("Shipping_instruction")) Then
                        Shippinginstruction = ""
                    Else
                        If (table2("Shipping_instruction") IsNot DBNull.Value) Then
                            Shippinginstruction = table2("Shipping_instruction").ToString.Replace("'", "")
                        Else
                            Shippinginstruction = ""
                        End If
                    End If
                    If (table2("shortage") > 0) Then
                        strkanban = String.Format("INSERT INTO T_KBPickingList (BuildPlanId,ItemNO,InventoryItemId,KBQty,MPQ,OnhandQty,AvailableQty,shortage,DeliveryDate,EJITID,MCAvailableQty,ShipInstruction)  values ({0},'{1}','{2}',{3},{4},{5},{6},{7},'{8}',{9},{10},'{11}')", BuildPlanId, table2("item_no"), table2("Inventory_item_id"), KBQty, MPQ, OnhandQty, table2("Available_Qty"), table2("shortage"), table2("Delivery_Date"), table2("ejit_id"), mcqty, Shippinginstruction)
                        da.ExecuteNonQuery(strkanban)
                    End If
                Next
                SaveData = True
            Catch ex As Exception
                ErrorLogging("KanBan-SaveBuildPlan", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                SaveData = False
            End Try
        End Using
    End Function




    Private Function isValidDJModel(ByVal dj_name As String, ByVal modelname As String, ByVal quantity As Decimal, ByVal wipQty As Decimal, ByVal OracleERPLogin As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            isValidDJModel = ""
            Dim ds As DataSet = New DataSet
            Dim oda As OracleDataAdapter = da.Oda_Sele
            Dim comm_submit As New OracleCommand()

            Try


                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.xxetr_kanban_pkg.match_dj_model"

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_success_flag", OracleType.VarChar, 10))
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_openqty", OracleType.VarChar, 2000))
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_error_mssg", OracleType.VarChar, 2000))

                oda.SelectCommand.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                oda.SelectCommand.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput
                oda.SelectCommand.Parameters("o_openqty").Direction = ParameterDirection.InputOutput

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OracleERPLogin.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_dj_name", OracleType.VarChar, 50)).Value = dj_name
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_model", OracleType.VarChar, 50)).Value = modelname

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "isValidDJ")
                oda.SelectCommand.Connection.Close()
                Dim Flag, Msg As String
                Dim openqty As Decimal

                Flag = oda.SelectCommand.Parameters("o_success_flag").Value.ToString
                Msg = oda.SelectCommand.Parameters("o_error_mssg").Value.ToString
                openqty = Decimal.Parse(oda.SelectCommand.Parameters("o_openqty").Value.ToString)

                If (Flag = "Y") Then
                    If (quantity > openqty) Then
                        isValidDJModel = "quantity should less than open quantity: " + openqty.ToString
                    End If
                    If (wipQty > openqty) Then
                        isValidDJModel = "WipQty should less than open quantity: " + openqty.ToString
                    End If
                Else
                    'isValidDJModel = "Error: Model and DJNo are not match!"
                    isValidDJModel = "Error: " & Msg
                End If


            Catch oe As Exception
                ErrorLogging("Kanban-isValidDJModel", OracleERPLogin.User, oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function


    Private Function isValidLocSubInv(ByVal subinventory As String, ByVal loc As String, ByVal OracleERPLogin As ERPLogin) As Boolean

        Using da As DataAccess = GetDataAccess()
            isValidLocSubInv = False
            Dim ds As DataSet = New DataSet
            Dim oda As OracleDataAdapter = da.Oda_Sele
            Dim comm_submit As New OracleCommand()

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.xxetr_wip_pkg.valid_subinvloc"

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_flag", OracleType.VarChar, 10))
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_msg", OracleType.VarChar, 2000))

                oda.SelectCommand.Parameters("o_flag").Direction = ParameterDirection.InputOutput
                oda.SelectCommand.Parameters("o_msg").Direction = ParameterDirection.InputOutput

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = OracleERPLogin.OrgCode
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_subinv", OracleType.VarChar, 50)).Value = subinventory
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_locator", OracleType.VarChar, 50)).Value = loc

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()
                Dim Flag, Msg As String

                Flag = oda.SelectCommand.Parameters("o_flag").Value.ToString
                Msg = oda.SelectCommand.Parameters("o_msg").Value.ToString


                If (Flag = "Y") Then
                    isValidLocSubInv = True
                End If


            Catch oe As Exception
                ErrorLogging("Kanban-isValidLocSubInv", OracleERPLogin.User, oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function GetNextBPID(ByVal OracleLoginData As ERPLogin) As String

        Dim NextKanbanID As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try

            Dim ra, k As Integer

            myConn.Open()
            NextKanbanID = ""
            k = 0
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = TimeOut_M5
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "BPID")

            While (k < 5 And NextKanbanID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextKanbanID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("GetNextBPID", "Deadlocked? " & Str(k), "Failed to get next BuildPlanID, " & ex.Message & ex.Source, "E")
                End Try
            End While
            myConn.Close()
            Return NextKanbanID
        Catch ex As Exception

            ErrorLogging("GetNextBPID", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return NextKanbanID
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function


    Public Function getBuildPlanList(ByVal startDate As String, ByVal endDate As String, ByVal productionfloor As String, ByVal p_ds As DataSet, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow
            Dim result As DataSet
            Try

                strsql = String.Format("select DISTINCT T_BuildPlanHeader.BuildplanId,createon from T_BuildPlanHeader,T_BuildPlanItem where  createon between  '{0}'and '{1}'  and productionfloor='{2}' and T_BuildPlanHeader.BuildPlanId=T_BuildPlanItem.BuildPlanId and T_BuildPlanItem.orgcode='{3}'", startDate, endDate, productionfloor, LoginData.OrgCode)
                result = da.ExecuteDataSet(strsql)

                If (result Is Nothing) Then
                    myDataRow = p_ds.Tables("BuildPlan").NewRow()
                    myDataRow("BuildplanId") = -1
                    myDataRow("createon") = "No record"
                    p_ds.Tables(0).Rows.Add(myDataRow)
                Else
                    myDataRow = p_ds.Tables("BuildPlan").NewRow()
                    myDataRow("BuildplanId") = -1
                    myDataRow("createon") = "please select..."
                    p_ds.Tables(0).Rows.Add(myDataRow)
                    For Each dr As DataRow In result.Tables(0).Rows
                        myDataRow = p_ds.Tables("BuildPlan").NewRow()
                        myDataRow("BuildplanId") = dr(0).ToString
                        myDataRow("createon") = dr(1).ToString
                        p_ds.Tables("BuildPlan").Rows.Add(myDataRow)
                    Next

                End If


                getBuildPlanList = p_ds


            Catch ex As Exception
                ErrorLogging("KanBan-getBuildPlanList", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function getBuildPlanDetail(ByVal buildPlanId As String, ByVal ischecked As Boolean, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow

            ' Dim myDataRow As Data.DataRow
            Dim result As DataTable
            getBuildPlanDetail = New DataSet
            Try

                strsql = String.Format("select createby,DJNO,model,subinventory,deliverydate,requireqty,orgcode from T_BuildPlanItem ,T_BuildPlanHeader where T_BuildPlanHeader. buildPlanId= T_BuildPlanItem. buildPlanId and T_BuildPlanItem. buildPlanId= {0}", buildPlanId)
                result = da.ExecuteDataTable(strsql)
                getBuildPlanDetail.Tables.Add(result)

                strsql = "select ItemNo,kbqty as requireQty,shortage,deliverydate,ordernumber,recqty,recon,recby,ejitid,ShipInstruction from T_KBPickinglist where buildPlanId= {0}"
                If (ischecked) Then
                    strsql = strsql + " and shortage<> recqty "
                End If
                strsql = String.Format(strsql, buildPlanId)
                result = da.ExecuteDataTable(strsql)
                getBuildPlanDetail.Tables.Add(result)

            Catch ex As Exception
                ErrorLogging("KanBan-getBuildPlanDetail", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function


    Public Function getLocatorsPB(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Dim myDataRow As Data.DataRow
            Dim result As DataSet
            Try

                strsql = String.Format("select DISTINCT Productionfloor from T_BuildPlanItem order by Productionfloor")
                getLocatorsPB = da.ExecuteDataSet(strsql)

            Catch ex As Exception
                ErrorLogging("KanBan-getLocatorsPB", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function getNow() As Date
        Using da As DataAccess = GetDataAccess()
            Dim strsql As String
            Try

                strsql = String.Format("Select   GetDate()")
                getNow = da.ExecuteScalar(strsql)

            Catch ex As Exception
                ErrorLogging("KanBan-getLocatorsPB", "", ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function


    Public Function getPONumber(ByVal buildplanDetail As DataSet, ByVal OracleERPLogin As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim temp As DataSet = New DataSet
            Dim strsql As String
            Dim i As Integer
            getPONumber = False
            Dim dr As DataRow() = buildplanDetail.Tables(1).Select("ordernumber is  null and shortage<>0")
            Dim oda As OracleDataAdapter = da.Oda_Sele
            Try
                oda.SelectCommand.Connection.Open()
                For i = 0 To dr.Length - 1
                    'If (Decimal.Parse(dr("recqty")) = 0) Then
                    oda.SelectCommand.CommandType = CommandType.StoredProcedure
                    oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_req_num"
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_ej_id", OracleType.VarChar, 25))
                    oda.SelectCommand.Parameters("p_ej_id").Value = dr(i)("ejitid").ToString
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_date", OracleType.VarChar, 25))
                    oda.SelectCommand.Parameters("p_date").Value = dr(i)("deliverydate").ToString
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_num", OracleType.VarChar, 240))
                    oda.SelectCommand.Parameters("o_num").Direction = ParameterDirection.Output
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_flag", OracleType.VarChar, 50))
                    oda.SelectCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                    oda.Fill(temp)
                    dr(i)("ordernumber") = oda.SelectCommand.Parameters("o_num").Value.ToString()
                    If (oda.SelectCommand.Parameters("o_flag").Value.ToString() = "Y") Then
                        strsql = String.Format("update T_Kbpickinglist set OrderNumber='{0}' where ejitid={1}", dr(i)("ordernumber"), dr(i)("ejitid").ToString)
                        da.ExecuteNonQuery(strsql)
                    End If
                    '  End If
                Next
                getPONumber = True
                oda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("KanBan-getPONumber", OracleERPLogin.User, ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function
    Public Function getPONumber() As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim temp As DataSet = New DataSet
            Dim strsql As String
            Dim i As Integer
            getPONumber = False
            Dim sqleTrace As String
            Dim TableResult As DataSet
            sqleTrace = String.Format("select * from T_kbpickinglist where ordernumber is  null")

            TableResult = da.ExecuteDataSet(sqleTrace)


            Dim dr As DataRow() = TableResult.Tables(0).Select("ordernumber is  null")
            Dim oda As OracleDataAdapter = da.Oda_Sele






            Try
                oda.SelectCommand.Connection.Open()
                For i = 0 To dr.Length - 1
                    'If (Decimal.Parse(dr("recqty")) = 0) Then
                    oda.SelectCommand.CommandType = CommandType.StoredProcedure
                    oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_req_num"
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_ej_id", OracleType.VarChar, 25))
                    oda.SelectCommand.Parameters("p_ej_id").Value = dr(i)("ejitid").ToString
                    oda.SelectCommand.Parameters.Add(New OracleParameter("p_date", OracleType.VarChar, 25))
                    oda.SelectCommand.Parameters("p_date").Value = dr(i)("deliverydate").ToString
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_num", OracleType.VarChar, 240))
                    oda.SelectCommand.Parameters("o_num").Direction = ParameterDirection.Output
                    oda.SelectCommand.Parameters.Add(New OracleParameter("o_flag", OracleType.VarChar, 50))
                    oda.SelectCommand.Parameters("o_flag").Direction = ParameterDirection.Output
                    oda.Fill(temp)
                    dr(i)("ordernumber") = oda.SelectCommand.Parameters("o_num").Value.ToString()
                    If (oda.SelectCommand.Parameters("o_flag").Value.ToString() = "Y") Then
                        strsql = String.Format("update T_Kbpickinglist set OrderNumber='{0}' where ejitid={1}", dr(i)("ordernumber"), dr(i)("ejitid").ToString)
                        da.ExecuteNonQuery(strsql)
                    End If
                    '  End If
                Next
                getPONumber = True
                oda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("KanBan-getPONumber", "test", ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

#End Region


#Region "upload"

    Public Function checkQSFormat(ByVal MRListData As DataSet, ByVal exceldata As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet

        Dim myDataRow As Data.DataRow
        Dim Productionfloor As String = ""
        Dim r1, r2 As DataRow()


        Using da As DataAccess = GetDataAccess()
            Try
                For Each dr As DataRow In exceldata.Tables(0).Rows
                    myDataRow = MRListData.Tables(0).NewRow()
                    myDataRow("action") = IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim)
                    myDataRow("ORIGINAL_ORG") = IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim)
                    myDataRow("ORG_CODE") = IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim)
                    myDataRow("SUPPLIER") = IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim)
                    myDataRow("SUPPLIER_SITE") = IIf(dr(6) Is DBNull.Value, "", dr(6).ToString().Trim)

                    myDataRow("ITEM_NO") = IIf(dr(9) Is DBNull.Value, "", dr(9).ToString.Trim)
                    myDataRow("MON") = IIf(dr(10) Is DBNull.Value, "", dr(10).ToString.Trim)
                    myDataRow("TUE") = IIf(dr(11) Is DBNull.Value, "", dr(11).ToString.Trim)
                    myDataRow("WED") = IIf(dr(12) Is DBNull.Value, "", dr(12).ToString.Trim)
                    myDataRow("TUR") = IIf(dr(13) Is DBNull.Value, "", dr(13).ToString.Trim)
                    myDataRow("FRI") = IIf(dr(14) Is DBNull.Value, "", dr(14).ToString().Trim)
                    myDataRow("SAT") = IIf(dr(15) Is DBNull.Value, "", dr(15).ToString.Trim)
                    myDataRow("SUN") = IIf(dr(16) Is DBNull.Value, "", dr(16).ToString.Trim)
                    myDataRow("ORIGINAL_LOCATION_id") = dr(21)
                    myDataRow("CHARGE_ACCOUNT_ID") = dr(23)

                    myDataRow("ORIGINAL_LOCATION") = IIf(dr(22) Is DBNull.Value, "", dr(22).ToString.Trim)
                    myDataRow("buyer_name") = IIf(dr(24) Is DBNull.Value, "", dr(24).ToString().Trim)


                    r1 = MRListData.Tables("errordata").Select("ORIGINAL_ORG='" & myDataRow("ORIGINAL_ORG") & "' and ORG_CODE='" & myDataRow("ORG_CODE") & "' and SUPPLIER='" & myDataRow("SUPPLIER").ToString.Replace("'", "''") & "' and SUPPLIER_SITE='" & myDataRow("SUPPLIER_SITE") & "' and ITEM_NO='" & myDataRow("ITEM_NO") & "'")
                    If (r1.Length > 0) Then
                        If (r1.Length = 1) Then
                            r1(0)("ErrorMsg") = "duplicate data!"
                        End If
                        myDataRow("ErrorMsg") = "duplicate data!"
                        MRListData.Tables(0).Rows.Add(myDataRow)
                        Continue For
                    End If


                    r2 = MRListData.Tables("errordata").Select(" ORG_CODE='" & myDataRow("ORG_CODE") & "' and ITEM_NO='" & myDataRow("ITEM_NO") & "' and ((MON='" & myDataRow("MON") & "' and MON='Y')  Or (TUE='" & myDataRow("TUE") & "' and TUE='Y')  or (WED='" & myDataRow("WED") & "' and WED='Y') or (TUR='" & myDataRow("TUR") & "' and TUR='Y') or (FRI='" & myDataRow("FRI") & "' and FRI='Y') or (SAT='" & myDataRow("SAT") & "' and SAT='Y') or (SUN='" & myDataRow("SUN") & "' and SUN='Y'))  and action<>'DEL'")
                    If (r2.Length > 0) Then
                        If (r2.Length = 1) Then
                            r2(0)("ErrorMsg") = "Flag 'Y' in the same date!"
                        End If
                        myDataRow("ErrorMsg") = "Flag 'Y' in the same date!"
                        MRListData.Tables(0).Rows.Add(myDataRow)
                        Continue For
                    End If

                    MRListData.Tables(0).Rows.Add(myDataRow)
                Next
                If (MRListData.Tables(0).Select("errormsg<>''").Length <> 0) Then
                    Return MRListData
                End If

            Catch ex As Exception
                ErrorLogging("MMC-checkTMFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try


            Dim i As Integer
            Dim Oda As OracleCommand = da.Oda_Insert.InsertCommand
            Try
                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_kanban_pkg.valid_quota_sharing"
                Oda.Parameters.Add("p_login_org_code", OracleType.VarChar, 255)
                Oda.Parameters.Add("P_ACTION", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_ORIGINAL_ORG", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ORG_CODE", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_SUPPLIER", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_SUPPLIER_SITE", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ITEM_NO", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_MON", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_TUE", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_WED", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_TUR", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_FRI", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_SAT", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_SUN", OracleType.VarChar, 50)
                Oda.Parameters.Add("p_original_location", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ORIGINAL_LOCATION_id", OracleType.Int32)
                Oda.Parameters.Add("P_CHARGE_ACCOUNT_ID", OracleType.Int32)
                Oda.Parameters.Add("P_Buyer_Name", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_message", OracleType.VarChar, 1240)


                For i = 0 To exceldata.Tables(0).Rows.Count - 1

                    Oda.Parameters("p_login_org_code").Value = OracleERPLogin.OrgCode
                    Oda.Parameters("P_ACTION").Value = exceldata.Tables(0).Rows(i)("ACTION")
                    Oda.Parameters("P_ORIGINAL_ORG").Value = exceldata.Tables(0).Rows(i)("ORIGINAL_ORG")
                    Oda.Parameters("P_ORG_CODE").Value = exceldata.Tables(0).Rows(i)("ORG_CODE")
                    Oda.Parameters("P_SUPPLIER").Value = String.Format(exceldata.Tables(0).Rows(i)("SUPPLIER"))
                    Oda.Parameters("P_SUPPLIER_SITE").Value = exceldata.Tables(0).Rows(i)("SUPPLIER_SITE")
                    Oda.Parameters("P_ITEM_NO").Value = exceldata.Tables(0).Rows(i)("ITEM_NO")
                    Oda.Parameters("P_MON").Value = exceldata.Tables(0).Rows(i)("MON")
                    Oda.Parameters("P_TUE").Value = exceldata.Tables(0).Rows(i)("TUE")
                    Oda.Parameters("P_WED").Value = exceldata.Tables(0).Rows(i)("WED")
                    Oda.Parameters("P_TUR").Value = exceldata.Tables(0).Rows(i)("TUR")
                    Oda.Parameters("P_FRI").Value = exceldata.Tables(0).Rows(i)("FRI")
                    Oda.Parameters("P_SAT").Value = exceldata.Tables(0).Rows(i)("SAT")
                    Oda.Parameters("P_SUN").Value = exceldata.Tables(0).Rows(i)("SUN")
                    Oda.Parameters("P_Buyer_Name").Value = exceldata.Tables(0).Rows(i)("Buyer_Name")
                    Oda.Parameters("p_original_location").Value = exceldata.Tables(0).Rows(i)("original_location")
                    Oda.Parameters("P_ORIGINAL_LOCATION_id").Value = exceldata.Tables(0).Rows(i)("ORIGINAL_LOCATION_id")
                    Oda.Parameters("P_CHARGE_ACCOUNT_ID").Value = exceldata.Tables(0).Rows(i)("CHARGE_ACCOUNT_ID")
                    Oda.Parameters("o_message").Direction = ParameterDirection.Output
                    Oda.ExecuteNonQuery()
                    MRListData.Tables(0).Rows(i)("errormsg") = Oda.Parameters("o_message").Value.ToString()
                Next
                Oda.Connection.Close()
                Return MRListData
            Catch oe As Exception
                ErrorLogging("Kanban-CheckQSFormat", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try

        End Using
    End Function

    Public Function uploadQuotaSharing(ByVal exceldata As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            uploadQuotaSharing = ""
            Dim i As Integer
            Dim Oda As OracleCommand = da.Oda_Insert.InsertCommand
            Try
                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "APPS.XXETR_KanBan_pkg.ins_quota_sharing"
                Oda.Parameters.Add("P_ACTION", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_ORIGINAL_ORG", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ORG_CODE", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_SUPPLIER", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_SUPPLIER_SITE", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ITEM_NO", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_MON", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_TUE", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_WED", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_TUR", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_FRI", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_SAT", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_SUN", OracleType.VarChar, 50)
                Oda.Parameters.Add("p_original_location", OracleType.VarChar, 500)
                Oda.Parameters.Add("P_ORIGINAL_LOCATION_id", OracleType.Int32)
                Oda.Parameters.Add("P_Buyer_Name", OracleType.VarChar, 50)
                Oda.Parameters.Add("P_CHARGE_ACCOUNT_ID", OracleType.Int32)
                Oda.Parameters.Add("o_message", OracleType.VarChar, 1240)


                For i = 0 To exceldata.Tables(0).Rows.Count - 1
                    Oda.Parameters("P_ACTION").Value = exceldata.Tables(0).Rows(i)("ACTION")
                    Oda.Parameters("P_ORIGINAL_ORG").Value = exceldata.Tables(0).Rows(i)("ORIGINAL_ORG")
                    Oda.Parameters("P_ORG_CODE").Value = exceldata.Tables(0).Rows(i)("ORG_CODE")
                    Oda.Parameters("P_SUPPLIER").Value = String.Format(exceldata.Tables(0).Rows(i)("SUPPLIER"))
                    Oda.Parameters("P_SUPPLIER_SITE").Value = exceldata.Tables(0).Rows(i)("SUPPLIER_SITE")
                    Oda.Parameters("P_ITEM_NO").Value = exceldata.Tables(0).Rows(i)("ITEM_NO")
                    Oda.Parameters("P_MON").Value = exceldata.Tables(0).Rows(i)("MON")
                    Oda.Parameters("P_TUE").Value = exceldata.Tables(0).Rows(i)("TUE")
                    Oda.Parameters("P_WED").Value = exceldata.Tables(0).Rows(i)("WED")
                    Oda.Parameters("P_TUR").Value = exceldata.Tables(0).Rows(i)("TUR")
                    Oda.Parameters("P_FRI").Value = exceldata.Tables(0).Rows(i)("FRI")
                    Oda.Parameters("P_SAT").Value = exceldata.Tables(0).Rows(i)("SAT")
                    Oda.Parameters("P_SUN").Value = exceldata.Tables(0).Rows(i)("SUN")
                    Oda.Parameters("P_Buyer_Name").Value = exceldata.Tables(0).Rows(i)("Buyer_Name")
                    Oda.Parameters("p_original_location").Value = exceldata.Tables(0).Rows(i)("original_location")
                    Oda.Parameters("P_ORIGINAL_LOCATION_ID").Value = exceldata.Tables(0).Rows(i)("ORIGINAL_LOCATION_ID")
                    Oda.Parameters("P_CHARGE_ACCOUNT_ID").Value = exceldata.Tables(0).Rows(i)("CHARGE_ACCOUNT_ID")
                    Oda.Parameters("o_message").Direction = ParameterDirection.Output
                    Oda.ExecuteNonQuery()
                    uploadQuotaSharing = Oda.Parameters("o_message").Value.ToString()
                Next
                Oda.Connection.Close()

            Catch oe As Exception

                ErrorLogging("Kanban-uploadQuotaSharing", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using

    End Function

    Public Function checkTMFormat(ByVal MRListData As DataSet, ByVal ExcelMRData As DataSet, ByVal OracleERPLogin As ERPLogin) As DataSet

        Dim myDataRow As Data.DataRow
        Dim Productionfloor As String = ""
        Dim r1 As DataRow()


        Using da As DataAccess = GetDataAccess()
            Try
                For Each dr As DataRow In ExcelMRData.Tables(0).Rows
                    myDataRow = MRListData.Tables(0).NewRow()
                    myDataRow("action") = IIf(dr(0) Is DBNull.Value, "", dr(0).ToString.Trim)
                    myDataRow("ou_name") = IIf(dr(1) Is DBNull.Value, "", dr(1).ToString.Trim)
                    myDataRow("vendor_name") = IIf(dr(2) Is DBNull.Value, "", dr(2).ToString.Trim)
                    myDataRow("system_time") = IIf(dr(3) Is DBNull.Value, "", dr(3).ToString.Trim)
                    myDataRow("new_system_time") = IIf(dr(4) Is DBNull.Value, "", dr(4).ToString.Trim)
                    myDataRow("local_time") = IIf(dr(5) Is DBNull.Value, "", dr(5).ToString().Trim)
                    myDataRow("new_local_time") = IIf(dr(6) Is DBNull.Value, "", dr(6).ToString().Trim)


                    If (myDataRow("action") = "DEL") Then
                        r1 = MRListData.Tables("errordata").Select("ou_name='" & myDataRow("ou_name") & "' and vendor_name='" & myDataRow("vendor_name") & "' and system_time='" & myDataRow("system_time") & "'")
                    Else
                        r1 = MRListData.Tables("errordata").Select("ou_name='" & myDataRow("ou_name") & "' and vendor_name='" & myDataRow("vendor_name") & "' and new_system_time='" & myDataRow("new_system_time") & "'")
                    End If
                    If (r1.Length > 0) Then
                        If (r1.Length = 1) Then
                            r1(0)("ErrorMsg") = "duplicate data!"
                        End If
                        myDataRow("ErrorMsg") = "duplicate data!"
                        MRListData.Tables(0).Rows.Add(myDataRow)
                        Continue For
                    End If

                    MRListData.Tables(0).Rows.Add(myDataRow)
                Next
                If (MRListData.Tables(0).Select("errormsg<>''").Length <> 0) Then
                    Return MRListData
                End If

            Catch ex As Exception
                ErrorLogging("MMC-checkTMFormat", OracleERPLogin.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try

            Dim Oda As OracleCommand = da.Oda_Insert.InsertCommand

            Try

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = " apps.xxetr_kanban_pkg.valid_vendor_deli_time"
                Oda.Parameters.Add("p_login_org_code", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_action", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_ou_name", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_vendor_name", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_system_time", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_new_system_time", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_local_time", OracleType.VarChar, 255)
                Oda.Parameters.Add("p_new_local_time", OracleType.VarChar, 255)
                Oda.Parameters.Add("o_message", OracleType.VarChar, 255)

                Dim i As Integer
                For i = 0 To MRListData.Tables(0).Rows.Count - 1
                    Oda.Parameters("p_login_org_code").Value = OracleERPLogin.OrgCode
                    Oda.Parameters("p_action").Value = MRListData.Tables(0).Rows(i)("action")
                    Oda.Parameters("p_ou_name").Value = MRListData.Tables(0).Rows(i)("ou_name")
                    Oda.Parameters("p_vendor_name").Value = MRListData.Tables(0).Rows(i)("vendor_name")
                    Oda.Parameters("p_system_time").Value = MRListData.Tables(0).Rows(i)("system_time")
                    Oda.Parameters("p_new_system_time").Value = MRListData.Tables(0).Rows(i)("new_system_time")
                    Oda.Parameters("p_local_time").Value = MRListData.Tables(0).Rows(i)("local_time")
                    Oda.Parameters("p_new_local_time").Value = MRListData.Tables(0).Rows(i)("new_local_time")
                    Oda.Parameters("o_message").Direction = ParameterDirection.InputOutput
                    Oda.ExecuteNonQuery()
                    MRListData.Tables(0).Rows(i)("errormsg") = Oda.Parameters("o_message").Value.ToString()
                    'comm.Parameters("o_message").SourceColumn = "message"
                Next
                Oda.Connection.Close()

                Return MRListData
            Catch oe As Exception
                ErrorLogging("MMC-checkTMFormat", OracleERPLogin.User.ToUpper, oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try

        End Using
    End Function
    Public Function uploadTransmission(ByVal exceldata As DataSet) As String
        'Using da As DataAccess = GetDataAccess()
        '    'Dim oda As OracleDataAdapter = New OracleDataAdapter()
        '    Dim comm As OracleCommand = da.Oda_Insert.InsertCommand
        '    uploadTransmission = ""
        '    Try
        '        'oda = da.Oda_Insert()


        Using da As DataAccess = GetDataAccess()
            uploadTransmission = ""
            Dim i As Integer
            Dim Oda As OracleCommand = da.Oda_Insert.InsertCommand
            Try
                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "APPS.XXETR_KanBan_pkg.ins_vendor_deli_time"
                Oda.Parameters.Add("p_action", OracleType.VarChar, 100)
                Oda.Parameters.Add("p_ou_name", OracleType.VarChar, 240)
                Oda.Parameters.Add("p_vendor_name", OracleType.VarChar, 240)
                Oda.Parameters.Add("p_system_time", OracleType.VarChar, 240)
                Oda.Parameters.Add("p_new_system_time", OracleType.VarChar, 240)
                Oda.Parameters.Add("p_local_time", OracleType.VarChar, 240)
                Oda.Parameters.Add("p_new_local_time", OracleType.VarChar, 240)
                Oda.Parameters.Add("o_message", OracleType.VarChar, 1240)

                'comm.Parameters("o_message").Direction = ParameterDirection.InputOutput

                'comm.Parameters("p_action").SourceColumn = "action"
                'comm.Parameters("p_ou_name").SourceColumn = "ou_name"
                'comm.Parameters("p_vendor_name").SourceColumn = "vendor_name"
                'comm.Parameters("p_system_time").SourceColumn = "system_time"
                'comm.Parameters("p_new_system_time").SourceColumn = "new_system_time"
                'comm.Parameters("p_local_time").SourceColumn = "local_time"
                'comm.Parameters("p_new_local_time").SourceColumn = "new_local_time"
                'comm.Parameters("o_message").SourceColumn = "message"
                'comm.Parameters("o_message").Direction = ParameterDirection.Output
                'uploadTransmission = comm.Parameters("o_message").Value.ToString()
                For i = 0 To exceldata.Tables(0).Rows.Count - 1
                    Oda.Parameters("p_action").Value = exceldata.Tables(0).Rows(i)("action")
                    Oda.Parameters("p_ou_name").Value = exceldata.Tables(0).Rows(i)("ou_name")
                    Oda.Parameters("p_vendor_name").Value = exceldata.Tables(0).Rows(i)("vendor_name")
                    Oda.Parameters("p_system_time").Value = exceldata.Tables(0).Rows(i)("system_time")
                    Oda.Parameters("p_new_system_time").Value = exceldata.Tables(0).Rows(i)("new_system_time")
                    Oda.Parameters("p_local_time").Value = exceldata.Tables(0).Rows(i)("local_time")
                    Oda.Parameters("p_new_local_time").Value = exceldata.Tables(0).Rows(i)("new_local_time")
                    Oda.Parameters("o_message").Direction = ParameterDirection.InputOutput
                    Oda.ExecuteNonQuery()
                    uploadTransmission = Oda.Parameters("o_message").Value.ToString()
                    'comm.Parameters("o_message").SourceColumn = "message"
                Next
                Oda.Connection.Close()
            Catch oe As Exception
                ErrorLogging("TDC-UploadTransmission", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using

    End Function
    Public Function getTransmission(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim transmission As DataSet = New DataSet
            Dim oda As OracleDataAdapter = New OracleDataAdapter
            Try
                oda = da.Oda_Sele()
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_transmission"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("p_org_code").Value = LoginData.OrgCode

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_cursor", OracleType.Cursor))
                oda.SelectCommand.Parameters.Add(New OracleParameter("o_message", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("o_cursor").Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters("o_message").Direction = ParameterDirection.Output


                oda.SelectCommand.Connection.Open()
                oda.Fill(transmission, "transmission")
                oda.SelectCommand.Connection.Close()


                Return transmission
            Catch oex As Exception
                ErrorLogging("Kanban-getTransmission", LoginData.User, oex.Message & oex.Source, "E")
                Throw oex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function


    Public Function getQuotaSharing(ByVal LoginData As ERPLogin, ByVal Item_no As String, ByVal buyername As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim QuotaSharing As DataSet = New DataSet
            Dim oda As OracleDataAdapter = New OracleDataAdapter
            Try
                oda = da.Oda_Sele()
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_KanBan_pkg.get_quota_sharing"
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("p_org_code").Value = LoginData.OrgCode

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_item_no", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("p_item_no").Value = Item_no

                oda.SelectCommand.Parameters.Add(New OracleParameter("p_buyer_Name", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("p_buyer_Name").Value = buyername

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_cursor", OracleType.Cursor))
                oda.SelectCommand.Parameters("o_cursor").Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(QuotaSharing, "QuotaSharing")
                oda.SelectCommand.Connection.Close()


                Return QuotaSharing
            Catch oex As Exception
                ErrorLogging("Kanban-getQuotaSharing", LoginData.User, oex.Message & oex.Source, "E")
                Throw oex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function checkPermission(ByVal LoginData As ERPLogin) As Boolean
        Dim flag As String
        Using da As DataAccess = GetDataAccess()

            Dim Oda As OracleCommand = da.OraCommand()

            Try

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_kanban_pkg.valid_upload_ejit"
                Oda.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240))
                Oda.Parameters("p_org_code").Value = LoginData.OrgCode

                Oda.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))
                Oda.Parameters("p_user_id").Value = LoginData.UserID
                Oda.Parameters.Add(New OracleParameter("o_flag", OracleType.VarChar, 240))
                Oda.Parameters("o_flag").Direction = ParameterDirection.Output
                Oda.Connection.Open()
                Oda.ExecuteNonQuery()

                flag = Oda.Parameters("o_flag").Value.ToString
                If (flag = "Y") Then
                    checkPermission = True
                Else
                    checkPermission = False
                End If
                Oda.Connection.Close()
            Catch oex As Exception
                checkPermission = False
                ErrorLogging("Kanban-checkPermission", LoginData.User, oex.Message & oex.Source, "E")
                Throw oex
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function getExceptionReport(ByVal LoginData As ERPLogin, ByVal report As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = New OracleDataAdapter
            Try
                oda = da.Oda_Sele()
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_kanban_pkg.get_flag_count "
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240))
                oda.SelectCommand.Parameters("p_org_code").Value = LoginData.OrgCode

                oda.SelectCommand.Parameters.Add(New OracleParameter("o_cursor", OracleType.Cursor))
                oda.SelectCommand.Parameters("o_cursor").Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(report, "exception")
                oda.SelectCommand.Connection.Close()


                Return report
            Catch oex As Exception
                ErrorLogging("Kanban-getExceptionReport", LoginData.User, oex.Message & oex.Source, "E")
                Throw oex
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

#End Region

End Class
