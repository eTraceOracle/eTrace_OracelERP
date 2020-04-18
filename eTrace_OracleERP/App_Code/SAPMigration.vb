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

Public Structure SAPPN_Check
    Public OraclePN As String
    Public Flag As Boolean
End Structure

Public Class SAPMigration
    Inherits PublicFunction

    'Public Function UpdateOraItem(ByVal Logindata As ERPLogin) As Boolean

    '    'Dim Logindata As ERPLogin = New ERPLogin
    '    'Logindata.User = "YUDY_LIU"
    '    'Logindata.OrgCode = "580"

    '    UpdateOraItem = True

    '    Dim da As DataAccess = GetDataAccess()
    '    Dim ds As DataSet = New DataSet

    '    Dim SAPPN, SAPUOM, OraItem, OraUOM As String
    '    Dim ConvFactor As Decimal
    '    Dim i As Integer
    '    Dim ra As Integer
    '    Dim DR(), drUOM() As DataRow

    '    Try
    '        ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step 'Update T_SAPPN' start")

    '        Dim Sqlstr As String
    '        Sqlstr = String.Format("Select * from T_SAPPN where OracleUOM IS NULL or OracleUOM = '' ")
    '        ds = da.ExecuteDataSet(Sqlstr, "SAPPN")

    '        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
    '        ElseIf ds.Tables(0).Rows.Count > 0 Then
    '            Dim dsUOMPart As DataSet = New DataSet
    '            Dim dsUOM As DataSet = New DataSet

    '            'Sqlstr = String.Format("Select * from T_SAPUOM_Part where Item = '{0}'", SAPPN)
    '            Sqlstr = String.Format("Select * from T_SAPUOM_Part")
    '            dsUOMPart = da.ExecuteDataSet(Sqlstr, "Items")

    '            'Sqlstr = String.Format("Select * from T_SAPUOM where SAPUOM = '{0}'", SAPUOM)
    '            Sqlstr = String.Format("Select * from T_SAPUOM ")
    '            dsUOM = da.ExecuteDataSet(Sqlstr, "UOM")

    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                SAPPN = ds.Tables(0).Rows(i)("SAPPN").ToString
    '                SAPUOM = ds.Tables(0).Rows(i)("SAPUOM").ToString

    '                DR = Nothing
    '                DR = dsUOMPart.Tables(0).Select(" SAPPN = '" & SAPPN & "'")
    '                If DR.Length > 0 Then
    '                    OraUOM = DR(0)("OracleUOM").ToString
    '                    ConvFactor = DR(0)("ConvFactor")
    '                Else
    '                    drUOM = Nothing
    '                    drUOM = dsUOM.Tables(0).Select(" SAPUOM = '" & SAPUOM & "'")
    '                    If drUOM.Length > 0 Then
    '                        OraUOM = drUOM(0)("OracleUOM").ToString
    '                        ConvFactor = drUOM(0)("ConvFactor")
    '                    Else
    '                        'No mapping UOM found, use SAPUOM instead
    '                        OraUOM = SAPUOM & "*"
    '                        ConvFactor = 1
    '                    End If
    '                End If

    '                ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set OracleUOM='{0}', ConvFactor='{1}' where SAPPN='{2}'", OraUOM, ConvFactor, SAPPN))
    '            Next

    '        End If


    '        ds = New DataSet
    '        Sqlstr = String.Format("Select * from T_SAPPN where Description IS NULL or Description = '' ")
    '        ds = da.ExecuteDataSet(Sqlstr, "OraItem")

    '        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
    '        ElseIf ds.Tables(0).Rows.Count > 0 Then

    '            Dim ItemData As DataSet = New DataSet()
    '            ItemData.Tables.Add("ItemData")
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("p_org_code", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_desc", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_commdity_code", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_rev", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_type_name", System.Type.GetType("System.String")))

    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_uom_code", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_routing_id", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_shelf_life_days", System.Type.GetType("System.Double")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_revision_control_code", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_lot_control_code", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_msl", System.Type.GetType("System.String")))

    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_valid_flag", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_revlist", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_sublist", System.Type.GetType("System.String")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_shelf_life_ctrl", System.Type.GetType("System.Double")))
    '            ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_rohs_class", System.Type.GetType("System.String")))

    '            Dim myDR As DataRow
    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                OraItem = ds.Tables(0).Rows(i)("OraclePN").ToString

    '                DR = Nothing
    '                DR = ItemData.Tables(0).Select(" p_item_num = '" & OraItem & "'")
    '                If DR.Length = 0 Then
    '                    myDR = ItemData.Tables(0).NewRow()
    '                    myDR("p_item_num") = OraItem
    '                    myDR("p_org_code") = Logindata.OrgCode
    '                    ItemData.Tables(0).Rows.Add(myDR)
    '                End If
    '            Next

    '            Dim dsItems As DataSet = New DataSet
    '            Try
    '                dsItems = get_iteminfo(ItemData, Logindata, "SAPMigration")
    '            Catch ex As Exception
    '                ErrorLogging("SAPMigration-UpdateOraItem-get_iteminfo", Logindata.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
    '            End Try

    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                Dim LotControl, ExpControl, RevControl As Boolean
    '                Dim Description, RoHS, SubInvList, RevList, Remarks As String

    '                SAPPN = ds.Tables(0).Rows(i)("SAPPN").ToString
    '                OraItem = ds.Tables(0).Rows(i)("OraclePN").ToString

    '                DR = Nothing
    '                DR = dsItems.Tables(0).Select(" p_item_num = '" & OraItem & "'")
    '                If DR.Length > 0 Then
    '                    OraUOM = DR(0)("o_uom_code").ToString
    '                    Description = DR(0)("o_item_desc").ToString
    '                    RoHS = DR(0)("o_rohs_class").ToString
    '                    SubInvList = DR(0)("o_sublist").ToString
    '                    RevList = DR(0)("o_revlist").ToString

    '                    If FixNull(DR(0)("o_lot_control_code")) = "" Then
    '                        Remarks = "LotControl error for Oracle Item " & OraItem
    '                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                        UpdateOraItem = False
    '                        Continue For
    '                    End If

    '                    If FixNull(DR(0)("o_shelf_life_ctrl")) = "" Then
    '                        Remarks = "ExpControl error for Oracle Item " & OraItem
    '                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                        UpdateOraItem = False
    '                        Continue For
    '                    End If

    '                    If FixNull(DR(0)("o_revision_control_code")) = "" Then
    '                        Remarks = "RevControl error for Oracle Item " & OraItem
    '                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                        UpdateOraItem = False
    '                        Continue For
    '                    End If

    '                    LotControl = False
    '                    If DR(0)("o_lot_control_code").ToString = "2" Then LotControl = True

    '                    ExpControl = False
    '                    If DR(0)("o_shelf_life_ctrl") <> 1 Then ExpControl = True

    '                    RevControl = False
    '                    If DR(0)("o_revision_control_code").ToString = "2" Then RevControl = True

    '                    Try
    '                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Description='{0}', RoHS='{1}', SubInvList='{2}', LotControl='{3}', ExpControl='{4}', " _
    '                           & "RevControl='{5}', RevList='{6}', Remarks='' where SAPPN='{7}'", Description, RoHS, SubInvList, LotControl, ExpControl, RevControl, RevList, SAPPN))
    '                        If ra < 1 Then
    '                            UpdateOraItem = False
    '                            Remarks = "Update error for Oracle Item " & OraItem
    '                            ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                        End If
    '                    Catch ex As Exception
    '                        UpdateOraItem = False
    '                        Remarks = "Update error for Oracle Item " & OraItem & ", " & ex.Message & ex.Source
    '                        ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Error: " & Remarks)
    '                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                    End Try
    '                Else
    '                    UpdateOraItem = False
    '                    Remarks = "No Data found for Oracle Item " & OraItem
    '                    ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
    '                End If
    '            Next

    '        End If

    '        'ds = New DataSet
    '        'Sqlstr = String.Format("Select * from T_SAPPN ")
    '        'ds = da.ExecuteDataSet(Sqlstr, "SAPPN")
    '        'Return ds
    '        If UpdateOraItem = False Then
    '            ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Error while updating table T_SAPPN, pls check for detail")
    '        End If
    '        ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step 'Update T_SAPPN' finish")

    '    Catch ex As Exception
    '        ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
    '        UpdateOraItem = False
    '        'Return Nothing
    '    End Try

    'End Function

    Public Function UpdateOraItem(ByVal Logindata As ERPLogin) As Boolean

        'Dim Logindata As ERPLogin = New ERPLogin
        'Logindata.User = "YUDY_LIU"
        'Logindata.OrgCode = "580"

        UpdateOraItem = True

        Dim da As DataAccess = GetDataAccess()
        Dim ds As DataSet = New DataSet

        Dim i, j As Integer
        Dim ra As Integer
        Dim SAPPN, SAPUOM, OraItem, OraUOM As String
        Dim ConvFactor As Decimal
        Dim DR(), drPart(), drUOM() As DataRow

        Try
            ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step: 'This batch Update T_SAPPN' start")

            Dim Sqlstr As String
            'Sqlstr = String.Format("Select * from T_SAPPN where OracleUOM IS NULL or OracleUOM = '' ")
            Sqlstr = String.Format("Select TOP (10000) SAPPN, OraclePN, SAPUOM, Description, ItemType, RoHS, OracleUOM, ConvFactor, SubInvList, LotControl, ExpControl, RevControl, RevList, Remarks from T_SAPPN where OracleUOM IS NULL or OracleUOM = '' ")
            ds = da.ExecuteDataSet(Sqlstr, "OraItem")

            If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "No record to be updated for 'T_SAPPN' ")
                Return True
            End If


            If ds.Tables(0).Rows.Count > 0 Then
                j = ds.Tables(0).Rows.Count
                'ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Total " & j & " records to be updated for 'T_SAPPN' ")
                ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "This batch " & j & " records to be updated for 'T_SAPPN' ")

                Dim dsUOMPart As DataSet = New DataSet
                Dim dsUOM As DataSet = New DataSet

                Sqlstr = String.Format("Select * from T_SAPUOM_Part")
                dsUOMPart = da.ExecuteDataSet(Sqlstr, "Items")

                Sqlstr = String.Format("Select * from T_SAPUOM ")
                dsUOM = da.ExecuteDataSet(Sqlstr, "UOM")

                Dim ItemData As DataSet = New DataSet()
                ItemData.Tables.Add("ItemData")
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("p_item_num", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("p_org_code", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_desc", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_commdity_code", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_rev", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_type_name", System.Type.GetType("System.String")))

                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_uom_code", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_routing_id", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_shelf_life_days", System.Type.GetType("System.Double")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_revision_control_code", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_lot_control_code", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_item_msl", System.Type.GetType("System.String")))

                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_valid_flag", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_revlist", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_sublist", System.Type.GetType("System.String")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_shelf_life_ctrl", System.Type.GetType("System.Double")))
                ItemData.Tables(0).Columns.Add(New Data.DataColumn("o_rohs_class", System.Type.GetType("System.String")))

                Dim myDR As DataRow
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    OraItem = ds.Tables(0).Rows(i)("OraclePN").ToString

                    DR = Nothing
                    DR = ItemData.Tables(0).Select("p_item_num = '" & OraItem & "'")
                    If DR.Length = 0 Then
                        myDR = ItemData.Tables(0).NewRow()
                        myDR("p_item_num") = OraItem
                        myDR("p_org_code") = Logindata.OrgCode
                        ItemData.Tables(0).Rows.Add(myDR)
                    End If
                Next

                'Sqlstr = String.Format("Select OraclePN as p_item_num, p_org_code='580', o_item_desc='', o_commdity_code='',o_item_rev='', o_type_name='', o_uom_code='',o_routing_id='', o_shelf_life_days='', " _
                '       & "o_revision_control_code='', o_lot_control_code='', o_item_msl='', o_valid_flag='',o_revlist='', o_sublist='', o_shelf_life_ctrl=0, o_rohs_class='' from T_SAPPN where OracleUOM IS NULL or OracleUOM = '' ")
                'ItemData = da.ExecuteDataSet(Sqlstr, "ItemData")


                ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step: 'Read Item Master from Oracle' start")

                Dim dsItems As DataSet = New DataSet
                Try
                    dsItems = get_iteminfo(ItemData, Logindata, "SAPMigration")
                Catch ex As Exception
                    ErrorLogging("SAPMigration-UpdateOraItem-get_iteminfo", Logindata.User.ToUpper, ex.Message & ex.Source, "E")
                End Try
                If dsItems Is Nothing OrElse dsItems.Tables.Count = 0 OrElse dsItems.Tables(0).Rows.Count = 0 Then
                    ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "No Data found from Oracle ")
                    Return False
                End If

                ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step: 'Read Item Master from Oracle' finished")


                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Dim LotControl, ExpControl, RevControl As Boolean
                    Dim Description, ItemType, RoHS, SubInvList, RevList, Remarks As String

                    SAPPN = ds.Tables(0).Rows(i)("SAPPN").ToString
                    SAPUOM = ds.Tables(0).Rows(i)("SAPUOM").ToString
                    OraItem = ds.Tables(0).Rows(i)("OraclePN").ToString

                    DR = Nothing
                    DR = dsItems.Tables(0).Select("p_item_num = '" & OraItem & "'")
                    If DR.Length = 0 Then
                        Remarks = "No Data found for Oracle Item " & OraItem
                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
                        UpdateOraItem = False
                        Continue For
                    End If


                    OraUOM = FixNull(DR(0)("o_uom_code"))
                    Description = FixNull(DR(0)("o_item_desc"))
                    ItemType = FixNull(DR(0)("o_type_name"))
                    RoHS = FixNull(DR(0)("o_rohs_class"))
                    SubInvList = FixNull(DR(0)("o_sublist"))
                    RevList = FixNull(DR(0)("o_revlist"))

                    Description = FilterSpecial(Description)

                    Dim ErrMsg As String = ""
                    If OraUOM = "" Then
                        ErrMsg = "OracleUOM" & "; "
                    End If

                    If ItemType = "" Then
                        ErrMsg = ErrMsg & "ItemType" & "; "
                    End If

                    'Get ConvFactor from table T_SAPUOM_Part or T_SAPUOM
                    drPart = Nothing
                    drPart = dsUOMPart.Tables(0).Select("SAPPN = '" & SAPPN & "' and SAPUOM = '" & SAPUOM & "' and OracleUOM = '" & OraUOM & "'")
                    If drPart.Length > 0 Then
                        'OraUOM = drPart(0)("OracleUOM").ToString
                        ConvFactor = drPart(0)("ConvFactor")
                    Else
                        drUOM = Nothing
                        drUOM = dsUOM.Tables(0).Select("SAPUOM = '" & SAPUOM & "' and OracleUOM = '" & OraUOM & "'")
                        If drUOM.Length > 0 Then
                            'OraUOM = drUOM(0)("OracleUOM").ToString
                            ConvFactor = drUOM(0)("ConvFactor")
                        Else
                            'No mapping ConvFactor, record error in Remarks
                            ErrMsg = ErrMsg & "ConvFactor" & "; "
                        End If
                    End If

                    If FixNull(DR(0)("o_lot_control_code")) = "" Then
                        ErrMsg = ErrMsg & "LotControl" & "; "
                    End If

                    If FixNull(DR(0)("o_shelf_life_ctrl")) = "" Then
                        ErrMsg = ErrMsg & "ExpControl" & "; "
                    End If

                    If FixNull(DR(0)("o_revision_control_code")) = "" Then
                        ErrMsg = ErrMsg & "RevControl" & "; "
                    End If

                    If ErrMsg <> "" Then
                        Remarks = "Error for Oracle Item " & OraItem & " with " & ErrMsg
                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
                        UpdateOraItem = False
                        Continue For
                    End If


                    LotControl = False
                    If DR(0)("o_lot_control_code").ToString = "2" Then LotControl = True

                    ExpControl = False
                    If DR(0)("o_shelf_life_ctrl") <> 1 Then ExpControl = True

                    RevControl = False
                    If DR(0)("o_revision_control_code").ToString = "2" Then RevControl = True

                    Try
                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Description='{0}', ItemType='{1}', RoHS='{2}', SubInvList='{3}', LotControl='{4}', ExpControl='{5}', RevControl='{6}', RevList='{7}', " _
                           & "OracleUOM='{8}', ConvFactor='{9}', Remarks='' where SAPPN='{10}'", Description, ItemType, RoHS, SubInvList, LotControl, ExpControl, RevControl, RevList, OraUOM, ConvFactor, SAPPN))
                        If ra < 1 Then
                            Remarks = "Update error for Oracle Item " & OraItem
                            ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
                            UpdateOraItem = False
                        End If
                    Catch ex As Exception
                        Remarks = "Update error for Oracle Item " & OraItem & ", " & ex.Message & ex.Source
                        ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, Remarks)
                        ra = da.ExecuteNonQuery(String.Format("update T_SAPPN set Remarks='{0}' where SAPPN='{1}'", Remarks, SAPPN))
                        UpdateOraItem = False
                    End Try
                Next

            End If

            'ds = New DataSet
            'Sqlstr = String.Format("Select * from T_SAPPN ")
            'ds = da.ExecuteDataSet(Sqlstr, "SAPPN")
            'Return ds

            ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, "Step 'This batch Update T_SAPPN' finish")

        Catch ex As Exception
            ErrorLogging("SAPMigration-UpdateOraItem", Logindata.User.ToUpper, ex.Message & ex.Source, "E")
            UpdateOraItem = False
            'Return Nothing
        End Try

    End Function

    Public Function ArchiveCLID(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("SAPMigration-ArchiveCLID", OracleLoginData.User.ToUpper, "Step 'ArchiveCLID' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction
            'Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_archive_clid_for_migration", myConn)
            'sda.SelectCommand.CommandType = CommandType.StoredProcedure

            'If sda.SelectCommand.ExecuteScalar.ToString = "0" Then
            '    Return False
            'ElseIf sda.SelectCommand.ExecuteScalar.ToString = "1" Then
            '    Return True
            'End If
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_archive_clid_for_migration", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("SAPMigration-ArchiveCLID", OracleLoginData.User.ToUpper, "Step 'ArchiveCLID' finish")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-ArchiveCLID", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function UpdateSTypeBin(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim strCMD As String

        ErrorLogging("SAPMigration-UpdateSTypeBin", OracleLoginData.User.ToUpper, "Step 'Initial SLOC/SType/SBin' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', MaterialRevision = '', LastTransaction = 'SAPIM_UpdateRevision' where MaterialRevision is NULL ", OracleLoginData.User)
            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageType = '', LastTransaction = 'SAPIM_UpdateSType' where StatusCode = 1 and StorageType is NULL ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-UpdateRevision", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            UpdateSTypeBin = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageType = NULL, LastTransaction = 'SAPIM_UpdateSType' where StorageType ='' ", OracleLoginData.User)
            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', SLOC = '', LastTransaction = 'SAPIM_UpdateSLOC' where SLOC is NULL ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-UpdateSLOC", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            UpdateSTypeBin = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageType = NULL, LastTransaction = 'SAPIM_UpdateSType' where StorageType ='' ", OracleLoginData.User)
            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageType = '', LastTransaction = 'SAPIM_UpdateSType' where StorageType is NULL ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-UpdateSType", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            UpdateSTypeBin = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        Try

            myConn.Open()

            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageBin = '', LastTransaction = 'SAPIM_UpdateSBin' where StorageBin is NULL ", OracleLoginData.User)
            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StorageType = '', LastTransaction = 'SAPIM_UpdateSType' where StatusCode = 1 and StorageType is NULL ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1800000
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-UpdateSBin", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            UpdateSTypeBin = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        UpdateSTypeBin = True
        ErrorLogging("SAPMigration-UpdateSTypeBin", OracleLoginData.User.ToUpper, "Step 'Initial SLOC/SType/SBin' finish")

    End Function

    Public Function DisableManualItems(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-DisableManualItems", OracleLoginData.User.ToUpper, "Step 'Disable Manual Items' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = '0', MigrationStatus = '9', LastTransaction = 'SAPIM_DisableSAPType' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin WHERE T_SAPSLOC.MigrationType = 'SAP' ", OracleLoginData.User)
            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = '0', MigrationStatus = '9', LastTransaction = 'SAPIM_DisableManualItems' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin WHERE T_SAPSLOC.MigrationType = 'SAP' AND StatusCode = '1'", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-DisableManualItems_SAPType", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            DisableManualItems = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            'strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = '0', MigrationStatus = '9', LastTransaction = 'SAPIM_DisableSAPType' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin WHERE T_SAPSLOC.MigrationType = 'SAP' ", OracleLoginData.User)
            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode = '0', MigrationStatus = '9', LastTransaction = 'SAPIM_DisableManualItems' FROM T_CLMaster LEFT OUTER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin WHERE StatusCode <> '0' AND (T_SAPSLOC.SLOC = '' OR T_SAPSLOC.SLOC IS NULL) AND (T_SAPSLOC.SType = '' OR T_SAPSLOC.SType IS NULL) AND (T_SAPSLOC.Bin = '' OR T_SAPSLOC.Bin IS NULL)", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-DisableManualItems_LocationNoMatch", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            DisableManualItems = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        DisableManualItems = True
        ErrorLogging("SAPMigration-DisableManualItems", OracleLoginData.User.ToUpper, "Step 'Disable Manual Items' finish")

    End Function

    Public Function CheckNoMapping(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        'ErrorLogging("SAPMigration-CheckMapping", OracleLoginData.User.ToUpper, "Step 'Check Mapping Records' start")
        'Try
        '    myConn.Open()
        '    SQLTransaction = myConn.BeginTransaction

        '    strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', MigrationStatus = NULL, ErrorMsg = '', LastTransaction = 'SAPIM_CheckMapping' WHERE (MaterialNo IN (SELECT SAPPN FROM T_SAPPN)) ", OracleLoginData.User)

        '    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
        '    CLMasterSQLCommand.Connection = myConn
        '    CLMasterSQLCommand.Transaction = SQLTransaction
        '    CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
        '    ra = CLMasterSQLCommand.ExecuteNonQuery()

        '    SQLTransaction.Commit()
        'Catch ex As Exception
        '    SQLTransaction.Rollback()
        '    ErrorLogging("SAPMigration-CheckMapping", "", "Error: " & ex.Message & ex.Source.ToString)
        '    CheckNoMapping = False
        '    Exit Function
        'Finally
        '    If myConn.State <> ConnectionState.Closed Then myConn.Close()
        'End Try
        'ErrorLogging("SAPMigration-CheckMapping", OracleLoginData.User.ToUpper, "Step 'Check Mapping Records' finish")

        ErrorLogging("SAPMigration-CheckNoPNMapping", OracleLoginData.User.ToUpper, "Step 'Check No PN Mapping Records' start")
        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("UPDATE T_CLMASTER set ChangedOn = getdate(), ChangedBy = '{0}', MigrationStatus = '0', ErrorMsg = 'No PN Mapping', LastTransaction = 'SAPIM_CheckNoPNMapping' WHERE (MaterialNo NOT IN (SELECT SAPPN FROM T_SAPPN WHERE NOT ConvFactor IS NULL)) ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-CheckNoPNMapping", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            CheckNoMapping = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        CheckNoMapping = True
        ErrorLogging("SAPMigration-CheckNoMapping", OracleLoginData.User.ToUpper, "Step 'Check No PN Mapping Records' finish")

    End Function

    Public Function CountNoMapping(ByVal OracleLoginData As ERPLogin) As Integer
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-CountNoPNMapping", OracleLoginData.User.ToUpper, "Step 'Count No PN Mapping Records' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("Select count(CLID) as Result from T_CLMASTER WHERE (StatusCode <> '0' and MigrationStatus = '0')")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            CountNoMapping = CLMasterSQLCommand.ExecuteScalar()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-CountNoPNMapping", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            CountNoMapping = Nothing
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        ErrorLogging("SAPMigration-CountNoPNMapping", OracleLoginData.User.ToUpper, "Step 'Count No PN Mapping Records' finish")

    End Function

    Public Function CheckMigrateStatus(ByVal OracleLoginData As ERPLogin) As Integer
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-CheckMigrateStatus", OracleLoginData.User.ToUpper, "Step 'Check MigrationStatus=0/NULL Records' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("Select count(CLID) as Result from T_CLMASTER WHERE (MigrationStatus = '0' OR MigrationStatus is NULL)")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            CheckMigrateStatus = CLMasterSQLCommand.ExecuteScalar()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-CheckMigrateStatus", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            CheckMigrateStatus = -1
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        ErrorLogging("SAPMigration-CheckMigrateStatus", OracleLoginData.User.ToUpper, "Step 'Check MigrationStatus=0/NULL Records' finish")

    End Function

    Public Function UpdateCLMaster(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-UpdateCLMaster", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("UPDATE T_CLMASTER set MaterialNo = T_SAPPN.OraclePN, MaterialDesc = T_SAPPN.Description, QtyBaseUOM = T_CLMaster.QtyBaseUOM * T_SAPPN.ConvFactor, BaseUOM = T_SAPPN.OracleUOM, ChangedOn = getdate(), ChangedBy = '{0}', ExpDate = CASE WHEN expcontrol = 'True' THEN (CASE WHEN not T_CLMaster.ExpDate is NULL THEN T_CLMaster.ExpDate  ELSE getdate() + 1095 END) ELSE T_CLMaster.ExpDate END, SLOC = T_SAPSLOC.SubInv, StorageType = '', StorageBin = T_SAPSLOC.Locator, OrgCode = '580', RoHS = T_SAPPN.RoHS, RTLot = CASE WHEN lotcontrol = 'True' THEN T_CLMaster.RecDocYear + T_CLMaster.RecDocNo ELSE '' END, SAPPN = T_CLMaster.MaterialNo, SAPQty = T_CLMaster.QtyBaseUOM, SAPUOM = T_CLMaster.BaseUOM, SAPSLOC = T_CLMaster.SLOC, SAPStorageType = T_CLMaster.StorageType, SAPStorageBin = T_CLMaster.StorageBin, MigrationStatus = '1', LastTransaction = 'SAPIM_UpdateCLMaster' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin INNER JOIN T_SAPPN ON T_CLMaster.MaterialNo = T_SAPPN.SAPPN  WHERE (T_CLMaster.MigrationStatus IS NULL) ", OracleLoginData.User)
            'strCMD = String.Format("UPDATE T_CLMASTER set MaterialNo = T_SAPPN.OraclePN, MaterialDesc = T_SAPPN.Description, QtyBaseUOM = T_CLMaster.QtyBaseUOM * T_SAPPN.ConvFactor, BaseUOM = T_SAPPN.OracleUOM, ChangedOn = getdate(), ChangedBy = '{0}', ExpDate = CASE WHEN expcontrol = 'True' THEN (CASE WHEN not T_CLMaster.ExpDate is NULL THEN T_CLMaster.ExpDate  ELSE getdate() + 1095 END) ELSE getdate() + 1095 END, SLOC = T_SAPSLOC.SubInv, StorageType = '', StorageBin = T_SAPSLOC.Locator, OrgCode = '580', RoHS = T_SAPPN.RoHS, RTLot = CASE WHEN lotcontrol = 'True' THEN T_CLMaster.RecDocYear + T_CLMaster.RecDocNo ELSE '' END, SAPPN = T_CLMaster.MaterialNo, SAPQty = T_CLMaster.QtyBaseUOM, SAPUOM = T_CLMaster.BaseUOM, SAPSLOC = T_CLMaster.SLOC, SAPStorageType = T_CLMaster.StorageType, SAPStorageBin = T_CLMaster.StorageBin, MigrationStatus = '1', LastTransaction = 'SAPIM_UpdateCLMaster' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SLOC = T_SAPSLOC.SLOC AND T_CLMaster.StorageType = T_SAPSLOC.SType AND T_CLMaster.StorageBin = T_SAPSLOC.Bin INNER JOIN T_SAPPN ON T_CLMaster.MaterialNo = T_SAPPN.SAPPN  WHERE (T_CLMaster.MigrationStatus IS NULL) ", OracleLoginData.User)
            'strCMD = String.Format("UPDATE T_CLMASTER set MaterialNo = T_SAPPN.OraclePN, MaterialDesc = T_SAPPN.Description, QtyBaseUOM = T_CLMaster.QtyBaseUOM * T_SAPPN.ConvFactor, BaseUOM = T_SAPPN.OracleUOM, ChangedOn = getdate(), ChangedBy = '{0}', ExpDate = CASE WHEN expcontrol = 'True' THEN T_CLMaster.ExpDate ELSE getdate() + 1095 END, SLOC = T_SAPSLOC.SubInv, StorageType = '', StorageBin = T_SAPSLOC.Locator, OrgCode = '580', RoHS = T_SAPPN.RoHS, RTLot = CASE WHEN lotcontrol = 'True' THEN T_CLMaster.RecDocYear + T_CLMaster.RecDocNo ELSE '' END, SAPPN = T_CLMaster.MaterialNo, SAPQty = T_CLMaster.QtyBaseUOM, SAPUOM = T_CLMaster.BaseUOM, SAPSLOC = T_CLMaster.SLOC, SAPStorageType = T_CLMaster.StorageType, SAPStorageBin = T_CLMaster.StorageBin, MigrationStatus = '1', LastTransaction = 'SAPIM_UpdateCLMaster' FROM T_CLMaster INNER JOIN T_SAPSLOC ON T_CLMaster.SAPSLOC = T_SAPSLOC.SLOC AND T_CLMaster.SAPStorageType = T_SAPSLOC.SType AND T_CLMaster.SAPStorageBin = T_SAPSLOC.Bin INNER JOIN T_SAPPN ON T_CLMaster.SAPPN = T_SAPPN.SAPPN  WHERE (T_CLMaster.MigrationStatus IS NULL) ", OracleLoginData.User)

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 360
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-UpdateCLMaster", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            UpdateCLMaster = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        UpdateCLMaster = True
        ErrorLogging("SAPMigration-UpdateCLMaster", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster' finish")

    End Function

    Public Function SumSAPIM(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-SumSAPIM", OracleLoginData.User.ToUpper, "Step 'Summarize to T_SAPIM' start")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("INSERT INTO T_SAPIM SELECT SAPPN, MaterialNo, MaterialRevision, SAPUOM, BaseUOM, SUM(SAPQty) AS SAPQty, SUM(QtyBaseUOM) AS QtyBaseUOM, UPPER(SAPSLOC) AS SAPSLOC, UPPER(SAPStorageType) AS SAPStorageType, UPPER(SAPStorageBin) AS SAPStorageBin, UPPER(SLOC) AS SLOC, UPPER(StorageBin) AS StorageBin, RTLot, ExpDate, RecDocNo, RecDocYear, '1' AS MigrationStatus, '' AS ClientID, '' AS TransactionID, '' AS ErrorMsg FROM T_CLMaster WHERE (StatusCode = 1) AND (SLOC <> '') AND (NOT (SLOC IS NULL)) AND (MigrationStatus = '1') GROUP BY SAPPN, MaterialNo, MaterialRevision, SAPUOM, BaseUOM, SAPSLOC, SAPStorageType, SAPStorageBin, SLOC, StorageBin, RTLot, ExpDate, RecDocNo, RecDocYear")
            'strCMD = String.Format("INSERT INTO T_SAPIM SELECT SAPPN, MaterialNo, MaterialRevision, SAPUOM, BaseUOM, SUM(SAPQty) AS SAPQty, SUM(QtyBaseUOM) AS QtyBaseUOM, SAPSLOC, SAPStorageType, SAPStorageBin, SLOC, StorageBin, RTLot, ExpDate, RecDocNo, RecDocYear, '1' AS MigrationStatus, '' as ClientID, '' AS TransactionID, '' AS ErrorMsg FROM T_CLMaster WHERE (StatusCode = 1) AND (SLOC <> '') AND (NOT (SLOC IS NULL)) AND MigrationStatus = '1' GROUP BY SAPPN, MaterialNo, MaterialRevision, SAPUOM, BaseUOM, SAPSLOC, SAPStorageType, SAPStorageBin, SLOC, StorageBin, RTLot, ExpDate, RecDocNo, RecDocYear")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 120
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-SumSAPIM", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            SumSAPIM = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        SumSAPIM = True
        ErrorLogging("SAPMigration-SumSAPIM", OracleLoginData.User.ToUpper, "Step 'Summarize to T_SAPIM' finish")

    End Function

    Public Function CheckQtyMatch(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ra As Integer
        Dim Sum1, Sum2 As Decimal
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-CheckQtyMatch", OracleLoginData.User.ToUpper, "Step 'Check if Qty Matches-T_CLMaster SumQty' start")
        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("Select sum(QtyBaseUOM) as Result from T_CLMASTER WHERE StatusCode = '1' and MigrationStatus = '1' ")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            Sum1 = CLMasterSQLCommand.ExecuteScalar()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-CheckQtyMatch-T_CLMaster", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            CheckQtyMatch = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        ErrorLogging("SAPMigration-CheckQtyMatch", OracleLoginData.User.ToUpper, "Step 'Check if Qty Matches-T_CLMaster SumQty' finish")

        ErrorLogging("SAPMigration-CheckQtyMatch", OracleLoginData.User.ToUpper, "Step 'Check if Qty Matches-T_SAPIM SumQty' start")
        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            strCMD = String.Format("Select sum(OracleQty) as Result from T_SAPIM WHERE MigrationStatus = '1' ")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
            Sum2 = CLMasterSQLCommand.ExecuteScalar()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("SAPMigration-CheckQtyMatch-T_SAPIM", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            CheckQtyMatch = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        ErrorLogging("SAPMigration-CheckQtyMatch-T_SAPIM", OracleLoginData.User.ToUpper, "Step 'Check if Qty Matches-T_SAPIM SumQty' finish")

        If Sum1 = Sum2 Then
            CheckQtyMatch = True
            ErrorLogging("SAPMigration-CheckQtyMatch-T_SAPIM", OracleLoginData.User.ToUpper, "Result for step 'Check if Qty Matches': SumQty in T_CLMaster " & Sum1 & " = SumQty in T_SAPIM " & Sum2)
        Else
            CheckQtyMatch = False
            ErrorLogging("SAPMigration-CheckQtyMatch-T_SAPIM", OracleLoginData.User.ToUpper, "Result for step 'Check if Qty Matches': SumQty in T_CLMaster " & Sum1 & " <> SumQty in T_SAPIM " & Sum2)
        End If

    End Function

    Public Function AssignClientID(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ClientID, ra, Cnt, Cnt_CLID As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        Try
            i = 0
            ClientID = 1
            'for WM
            ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' for WM start")
            Do
                Try
                    i = i + 1
                    If i > 10 Then
                        i = 1
                        ClientID = ClientID + 1
                    End If
                    Cnt = Cnt + 1

                    myConn.Open()
                    SQLTransaction = myConn.BeginTransaction

                    ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' round: " & Cnt)
                    strCMD = String.Format("UPDATE TOP(1000) T_SAPIM SET ClientID = '{0}' WHERE (Stype <> '') and (Bin <> '') and (ClientID is NULL or ClientID = '')", ClientID)

                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.Connection = myConn
                    CLMasterSQLCommand.Transaction = SQLTransaction
                    CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 120
                    ra = CLMasterSQLCommand.ExecuteNonQuery()

                    SQLTransaction.Commit()
                Catch ex As Exception
                    SQLTransaction.Rollback()
                    ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
                    AssignClientID = False
                    Exit Function
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try

                Using DA As DataAccess = GetDataAccess()
                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_SAPIM where (Stype <> '') and (Bin <> '') and (ClientID = '' OR ClientID is NULL)")
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Not Cnt_CLID > 0 Then
                        Exit Do
                    End If
                End Using
            Loop
            ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' for WM finish")

            ' for IM
            i = 0
            ClientID = ClientID + 1
            ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' for IM start")
            Do
                Try
                    i = i + 1
                    If i > 10 Then
                        i = 1
                        ClientID = ClientID + 1
                    End If
                    Cnt = Cnt + 1

                    myConn.Open()
                    SQLTransaction = myConn.BeginTransaction

                    ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' round: " & Cnt)
                    strCMD = String.Format("UPDATE TOP(1000) T_SAPIM SET ClientID = '{0}' WHERE (Stype = '') and (Bin = '') and (ClientID is NULL or ClientID = '')", ClientID)

                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.Connection = myConn
                    CLMasterSQLCommand.Transaction = SQLTransaction
                    CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 120
                    ra = CLMasterSQLCommand.ExecuteNonQuery()

                    SQLTransaction.Commit()
                Catch ex As Exception
                    SQLTransaction.Rollback()
                    ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
                    AssignClientID = False
                    Exit Function
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try

                Using DA As DataAccess = GetDataAccess()
                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_SAPIM where (Stype = '') and (Bin = '') and (ClientID = '' OR ClientID is NULL)")
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Not Cnt_CLID > 0 Then
                        Exit Do
                    End If
                End Using
            Loop
            ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Step 'Assign ClientID' for IM finish")

        Catch ex As Exception
            ErrorLogging("SAPMigration-AssignClientID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            AssignClientID = False
            Exit Function
        Finally
        End Try
        AssignClientID = True
    End Function

    Public Function UploadToOracle(ByVal ClientID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, j, cnt, Cnt_CLID, Err_Chk As Integer
        Dim TransactionID As Long
        Dim strCMD, strCMD1, strCMD2 As String
        Dim ds, ds_return, d_temp As DataSet
        Dim eTraceIM As eTraceWS.eTraceOracleERP = New eTraceWS.eTraceOracleERP
        Dim OracleLogin As eTraceWS.ERPLogin = New eTraceWS.ERPLogin
        Dim DR() As DataRow

        OracleLogin.Server = OracleLoginData.Server
        OracleLogin.User = OracleLoginData.User
        OracleLogin.PWD = OracleLoginData.PWD
        OracleLogin.OrgCode = OracleLoginData.OrgCode
        OracleLogin.Application = OracleLoginData.Application
        OracleLogin.UserID = OracleLoginData.UserID
        OracleLogin.UserType = OracleLoginData.UserType
        OracleLogin.UserDept = OracleLoginData.UserDept
        OracleLogin.ResetFlag = OracleLoginData.ResetFlag
        OracleLogin.Printer = OracleLoginData.Printer
        OracleLogin.ErrorMsg = OracleLoginData.ErrorMsg
        OracleLogin.AppID_PO = OracleLoginData.AppID_PO
        OracleLogin.RespID_PO = OracleLoginData.RespID_PO
        OracleLogin.AppID_Inv = OracleLoginData.AppID_Inv
        OracleLogin.RespID_Inv = OracleLoginData.RespID_Inv
        OracleLogin.AppID_WIP = OracleLoginData.AppID_WIP
        OracleLogin.RespID_WIP = OracleLoginData.RespID_WIP
        OracleLogin.AppID_KB = OracleLoginData.AppID_KB
        OracleLogin.RespID_KB = OracleLoginData.RespID_KB
        OracleLogin.ClientVersion = OracleLoginData.ClientVersion
        OracleLogin.ProductionLine = OracleLoginData.ProductionLine

        ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Step 'Load Records to Oracle' start")

        Try
            Using DA As DataAccess = GetDataAccess()
                Err_Chk = 0
                'TransactionID = OracleLoginData.OrgCode & OracleLoginData.UserID
                Do
                    ds = New DataSet
                    ds_return = New DataSet

                    cnt = cnt + 1
                    Err_Chk = 1
                    TransactionID = eTraceIM.GetNextInvID(OracleLogin)

                    Err_Chk = 2
                    d_temp = New DataSet
                    strCMD1 = String.Format("Select TOP 1000 ID, '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, (CASE WHEN ItemType = 'FG' THEN 'INVTY CONV - FG' ELSE 'INVTY CONV - RM/SA' END) as p_transaction_source_name, T_SAPIM.OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, T_SAPIM.OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_destination, Locator as p_locator_destination, RTLot as p_lot_number, ExpDate as p_lot_expiration_date, '' as p_reason_code, T_SAPIM.SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM inner join T_SAPPN on T_SAPIM.SAPPN = T_SAPPN.SAPPN WHERE ClientID = '{3}' and MigrationStatus = '1' ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    'strCMD1 = String.Format("Select TOP 1000 ID, '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, 'INVTY CONV - RM/SA' as p_transaction_source_name, OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_destination, Locator as p_locator_destination, RTLot as p_lot_number, ExpDate as p_lot_expiration_date, '' as p_reason_code, SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM WHERE ClientID = '{3}' and MigrationStatus = '1' ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    d_temp = DA.ExecuteDataSet(strCMD1, "accountreceipt_table")
                    ds = d_temp.Copy
                    ds.Tables(0).Columns.RemoveAt(0)
                    Err_Chk = 3
                    'strCMD1 = String.Format("Select TOP 1000 '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, 'INVTY CONV - RM/SA' as p_transaction_source_name, OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_destination, Locator as p_locator_destination, RTLot as p_lot_number, ExpDate as p_lot_expiration_date, '' as p_reason_code, SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM WHERE ClientID = '{3}' and MigrationStatus = '1' ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    'ds = DA.ExecuteDataSet(strCMD1, "accountreceipt_table")

                    For Each drow As DataRow In ds.Tables(0).Rows
                        drow.SetAdded()
                    Next
                    Err_Chk = 4
                    If ds.Tables(0).Rows.Count > 0 Then

                        ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Step 'Load Records to Oracle' round: " & cnt)
                        'eTraceIM.Timeout = 1000 * 60 * 60
                        ds_return = batch_receipt(ds, OracleLoginData, "SAPIM_UploadToOracle", TransactionID, "")
                        'ds_return = account_alias_batch_receipt(ds, OracleLoginData, "SAPMigration_UploadToOracle_account_alias_batch_receipt", TransactionID, "")
                    Else
                        Exit Do
                    End If
                    Err_Chk = 5
                    DR = Nothing
                    DR = ds_return.Tables("accountreceipt_table").Select("o_return_status = 'N'")

                    If DR.Length = 0 Then
                        ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " is successful in Oracle")
                        For i = 0 To d_temp.Tables("accountreceipt_table").Rows.Count - 1
                            strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '2', TransactionID = '{1}', ErrorMsg = ''  WHERE (ID = '{0}') ", d_temp.Tables("accountreceipt_table").Rows(i)("ID"), TransactionID)
                            DA.ExecuteNonQuery(strCMD2)
                        Next
                    Else
                        ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Error: Round" & cnt & " is failed in Oracle")
                        j = j + 1
                        For i = 0 To d_temp.Tables("accountreceipt_table").Rows.Count - 1
                            strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', TransactionID = '{2}', ErrorMsg = '{1}' WHERE (ID = '{0}') ", d_temp.Tables("accountreceipt_table").Rows(i)("ID"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("o_return_message")), TransactionID)
                            'strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountreceipt_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_item_revision")), ds_return.Tables("accountreceipt_table").Rows(i)("p_subinventory_destination"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_locator_destination")), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_lot_number")), "Post failed", ClientID)
                            DA.ExecuteNonQuery(strCMD2)
                        Next
                    End If
                    Err_Chk = 6
                    'If DR.Length = 0 Then
                    '    ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " is successful in Oracle")
                    '    For i = 0 To ds_return.Tables("accountreceipt_table").Rows.Count - 1
                    '        strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '2', ErrorMsg = ''  WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{5}') ", ds_return.Tables("accountreceipt_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_item_revision")), ds_return.Tables("accountreceipt_table").Rows(i)("p_subinventory_destination"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_locator_destination")), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_lot_number")), ClientID)
                    '        DA.ExecuteNonQuery(strCMD2)
                    '    Next
                    'Else
                    '    ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Error: Round" & cnt & " is failed in Oracle")
                    '    j = j + 1
                    '    For i = 0 To ds_return.Tables("accountreceipt_table").Rows.Count - 1
                    '        strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountreceipt_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_item_revision")), ds_return.Tables("accountreceipt_table").Rows(i)("p_subinventory_destination"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_locator_destination")), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_lot_number")), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("o_return_message")), ClientID)
                    '        'strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountreceipt_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_item_revision")), ds_return.Tables("accountreceipt_table").Rows(i)("p_subinventory_destination"), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_locator_destination")), FixNull(ds_return.Tables("accountreceipt_table").Rows(i)("p_lot_number")), "Post failed", ClientID)
                    '        DA.ExecuteNonQuery(strCMD2)
                    '    Next
                    'End If
                    ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " finish updating accordingly in eTrace table")

                    ds.Clear()
                    ds_return.Clear()
                    d_temp.Clear()
                    Err_Chk = 7
                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_SAPIM where ClientID = '{0}' and MigrationStatus = '1' ", ClientID)
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Cnt_CLID < 1 Then
                        Exit Do
                    End If
                    If cnt > 2000 Then
                        Exit Do
                    End If
                    Err_Chk = 8
                Loop
            End Using
            If j < 1 Then
                UploadToOracle = True
            Else
                UploadToOracle = False
            End If
            Err_Chk = 9
            ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Step 'Load Records to Oracle' finish")
        Catch ex As Exception
            ErrorLogging("SAPMigration-UploadToOracle", OracleLoginData.User.ToUpper, "Error: " & Err_Chk & "  " & ex.Message & ex.Source.ToString)
            UploadToOracle = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function IssueFmOracle(ByVal ClientID As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, j, cnt, Cnt_CLID As Integer
        Dim TransactionID As Long
        Dim strCMD, strCMD1, strCMD2 As String
        Dim ds, ds_return, d_temp As DataSet
        Dim eTraceIM As eTraceWS.eTraceOracleERP = New eTraceWS.eTraceOracleERP
        'Dim OracleLogin As eTraceWS.ERPLogin = New eTraceWS.ERPLogin
        Dim DR() As DataRow

        'OracleLogin.Server = OracleLoginData.Server
        'OracleLogin.User = OracleLoginData.User
        'OracleLogin.PWD = OracleLoginData.PWD
        'OracleLogin.OrgCode = OracleLoginData.OrgCode
        'OracleLogin.Application = OracleLoginData.Application
        'OracleLogin.UserID = OracleLoginData.UserID
        'OracleLogin.UserType = OracleLoginData.UserType
        'OracleLogin.UserDept = OracleLoginData.UserDept
        'OracleLogin.ResetFlag = OracleLoginData.ResetFlag
        'OracleLogin.Printer = OracleLoginData.Printer
        'OracleLogin.ErrorMsg = OracleLoginData.ErrorMsg
        'OracleLogin.AppID_PO = OracleLoginData.AppID_PO
        'OracleLogin.RespID_PO = OracleLoginData.RespID_PO
        'OracleLogin.AppID_Inv = OracleLoginData.AppID_Inv
        'OracleLogin.RespID_Inv = OracleLoginData.RespID_Inv
        'OracleLogin.AppID_WIP = OracleLoginData.AppID_WIP
        'OracleLogin.RespID_WIP = OracleLoginData.RespID_WIP
        'OracleLogin.AppID_KB = OracleLoginData.AppID_KB
        'OracleLogin.RespID_KB = OracleLoginData.RespID_KB
        'OracleLogin.ClientVersion = OracleLoginData.ClientVersion
        'OracleLogin.ProductionLine = OracleLoginData.ProductionLine

        ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Step 'Misc Issue from Oracle' start")

        Try
            Using DA As DataAccess = GetDataAccess()
                TransactionID = OracleLoginData.OrgCode & OracleLoginData.UserID
                Do
                    ds = New DataSet
                    ds_return = New DataSet
                    cnt = cnt + 1
                    TransactionID = TransactionID + 1

                    d_temp = New DataSet
                    strCMD1 = String.Format("Select TOP 1000 ID, '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, (CASE WHEN ItemType = 'FG' THEN 'INVTY CONV - FG' ELSE 'INVTY CONV - RM/SA' END) as p_transaction_source_name, T_SAPIM.OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, T_SAPIM.OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_source, Locator as p_locator_source, RTLot as p_lot_number, '' as p_reason_code, T_SAPIM.SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM inner join T_SAPPN on T_SAPIM.SAPPN = T_SAPPN.SAPPN WHERE ClientID = '{3}' and (MigrationStatus = '2' or MigrationStatus = '3') ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    'strCMD1 = String.Format("Select TOP 1000 ID, '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, 'INVTY CONV - RM/SA' as p_transaction_source_name, OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_source, Locator as p_locator_source, RTLot as p_lot_number, '' as p_reason_code, SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM WHERE ClientID = '{3}' and (MigrationStatus = '2' or MigrationStatus = '3') ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    d_temp = DA.ExecuteDataSet(strCMD1, "accountissue_table")
                    ds = d_temp.Copy
                    ds.Tables(0).Columns.RemoveAt(0)

                    'strCMD1 = String.Format("Select TOP 1000 '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, 'INVTY CONV - RM/SA' as p_transaction_source_name, OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_source, Locator as p_locator_source, RTLot as p_lot_number, '' as p_reason_code, SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM WHERE ClientID = '{3}' and (MigrationStatus = '2' or MigrationStatus = '3') ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID, ClientID)
                    ''strCMD1 = String.Format("Select TOP 1000 '1800000' as p_timeout, '{0}' as p_organization_code, '{1}' as p_transaction_header_id, 'INVTY CONV - RM/SA' as p_transaction_source_name, OracleUOM as p_transaction_uom, '99'  as p_source_line_id, '99' as p_source_header_id,'{2}' as p_user_id, OraclePN as p_item_segment1, Rev as p_item_revision, Subinv as p_subinventory_source, Locator as p_locator_source, RTLot as p_lot_number, '' as p_reason_code, SAPPN as p_transaction_reference, OracleQty as p_transaction_quantity, OracleQty as p_primary_quantity, '' as o_return_status, '' as o_return_message from T_SAPIM WHERE (MigrationStatus = '2' or MigrationStatus = '3') ", OracleLoginData.OrgCode, TransactionID, OracleLoginData.UserID)
                    'ds = DA.ExecuteDataSet(strCMD1, "accountissue_table")

                    For Each drow As DataRow In ds.Tables(0).Rows
                        drow.SetAdded()
                    Next

                    If ds.Tables(0).Rows.Count > 0 Then
                        ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Step 'Misc Issue from Oracle' round: " & cnt)
                        'eTraceIM.Timeout = 1000 * 60 * 60
                        ds_return = batch_issue(ds, OracleLoginData, "SAPIM_IssueFmOracle", TransactionID, "")
                        'ds_return = account_alias_batch_issue(ds, OracleLoginData, "SAPMigration_IssueFmOracle_account_alias_batch_issue", TransactionID, "")
                    Else
                        Exit Do
                    End If

                    DR = Nothing
                    DR = ds_return.Tables("accountissue_table").Select("o_return_status = 'N'")

                    If DR.Length = 0 Then
                        ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " is successful in Oracle")
                        For i = 0 To d_temp.Tables("accountissue_table").Rows.Count - 1
                            strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '4', ErrorMsg = ''  WHERE (ID = '{0}') ", d_temp.Tables("accountissue_table").Rows(i)("ID"))
                            DA.ExecuteNonQuery(strCMD2)
                        Next
                    Else
                        ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Error: Round" & cnt & " is failed in Oracle")
                        j = j + 1
                        For i = 0 To d_temp.Tables("accountissue_table").Rows.Count - 1
                            strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{1}' WHERE (ID = '{0}') ", d_temp.Tables("accountissue_table").Rows(i)("ID"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("o_return_message")))
                            'strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountissue_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_item_revision")), ds_return.Tables("accountissue_table").Rows(i)("p_subinventory_source"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_locator_source")), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_lot_number")), "Post failed", ClientID)
                            DA.ExecuteNonQuery(strCMD2)
                        Next
                    End If

                    'If DR.Length = 0 Then
                    '    ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " is successful in Oracle")
                    '    For i = 0 To ds_return.Tables("accountissue_table").Rows.Count - 1
                    '        strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '4', ErrorMsg = ''  WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{5}') ", ds_return.Tables("accountissue_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_item_revision")), ds_return.Tables("accountissue_table").Rows(i)("p_subinventory_source"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_locator_source")), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_lot_number")), ClientID)
                    '        DA.ExecuteNonQuery(strCMD2)
                    '    Next
                    'Else
                    '    ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Error: Round" & cnt & " is failed in Oracle")
                    '    j = j + 1
                    '    For i = 0 To ds_return.Tables("accountissue_table").Rows.Count - 1
                    '        strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountissue_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_item_revision")), ds_return.Tables("accountissue_table").Rows(i)("p_subinventory_source"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_locator_source")), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_lot_number")), FixNull(ds_return.Tables("accountissue_table").Rows(i)("o_return_message")), ClientID)
                    '        'strCMD2 = String.Format("UPDATE T_SAPIM set MigrationStatus = '0', ErrorMsg = '{5}' WHERE (OraclePN = '{0}' and Rev = '{1}' and Subinv = '{2}' and Locator = '{3}' and RTLot = '{4}' and ClientID = '{6}') ", ds_return.Tables("accountissue_table").Rows(i)("p_item_segment1"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_item_revision")), ds_return.Tables("accountissue_table").Rows(i)("p_subinventory_source"), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_locator_source")), FixNull(ds_return.Tables("accountissue_table").Rows(i)("p_lot_number")), "Post failed", ClientID)
                    '        DA.ExecuteNonQuery(strCMD2)
                    '    Next
                    'End If
                    ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Round: " & cnt & " finish updating accordingly in eTrace table")

                    ds.Clear()
                    ds_return.Clear()
                    d_temp.Clear()

                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_SAPIM where ClientID = '{0}' and (MigrationStatus = '2' or MigrationStatus = '3')", ClientID)
                    'strCMD = String.Format("Select count(*) from T_SAPIM where (MigrationStatus = '2' or MigrationStatus = '3')")
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Cnt_CLID < 1 Then
                        Exit Do
                    End If
                    If cnt > 2000 Then
                        Exit Do
                    End If
                Loop
            End Using
            If j < 1 Then
                IssueFmOracle = True
            Else
                IssueFmOracle = False
            End If
            ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Step 'Misc Issue from Oracle' finish")
        Catch ex As Exception
            ErrorLogging("SAPMigration-IssueFmOracle", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            IssueFmOracle = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function RollbackCLIDInfo(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim i, ClientID, ra, Cnt, Cnt_CLID As Integer
        Dim strCMD As String
        Dim ds As New DataSet

        ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Status' start")
        Try
            i = 0
            ClientID = 1
            Do
                Try
                    i = i + 1
                    If i > 10 Then
                        i = 1
                        ClientID = ClientID + 1
                    End If
                    Cnt = Cnt + 1

                    myConn.Open()
                    SQLTransaction = myConn.BeginTransaction

                    ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Status' round: " & Cnt)
                    strCMD = String.Format("UPDATE TOP(10000) T_CLMaster SET MigrationStatus = NULL, StatusCode = '1'  WHERE (CLID NOT LIKE 'ZSP%') AND (StatusCode <> '1') AND (SLOC <> '') AND (NOT (SLOC IS NULL))")

                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.Connection = myConn
                    CLMasterSQLCommand.Transaction = SQLTransaction
                    CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
                    ra = CLMasterSQLCommand.ExecuteNonQuery()

                    SQLTransaction.Commit()
                Catch ex As Exception
                    SQLTransaction.Rollback()
                    ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
                    RollbackCLIDInfo = False
                    Exit Function
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try

                Using DA As DataAccess = GetDataAccess()
                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_CLMaster where (CLID NOT LIKE 'ZSP%') AND (StatusCode <> '1') AND (SLOC <> '') AND (NOT (SLOC IS NULL))")
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Not Cnt_CLID > 0 Then
                        Exit Do
                    End If
                End Using
            Loop
            ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Status' finish")

            ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Mapping' start")
            i = 0
            ClientID = 1
            Do
                Try
                    i = i + 1
                    If i > 10 Then

                        i = 1
                        ClientID = ClientID + 1
                    End If
                    Cnt = Cnt + 1

                    myConn.Open()
                    SQLTransaction = myConn.BeginTransaction

                    ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Mapping' round: " & Cnt)
                    strCMD = String.Format("UPDATE TOP(10000) T_CLMaster SET MaterialNo = SAPPN, QtyBaseUOM = SAPQty, BaseUOM = SAPUOM, StatusCode = '1', SLOC = SAPSLOC, StorageType = SAPStorageType, StorageBin = SAPStorageBin, MigrationStatus = NULL, SAPPN = NULL, SAPQty = NULL, SAPUOM = NULL, SAPSLOC = NULL, SAPStorageType = NULL,SAPStorageBin = NULL WHERE (SAPPN <> '') AND (NOT (SAPPN IS NULL))")

                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.Connection = myConn
                    CLMasterSQLCommand.Transaction = SQLTransaction
                    CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 60
                    ra = CLMasterSQLCommand.ExecuteNonQuery()

                    SQLTransaction.Commit()
                Catch ex As Exception
                    SQLTransaction.Rollback()
                    ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
                    RollbackCLIDInfo = False
                    Exit Function
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try

                Using DA As DataAccess = GetDataAccess()
                    Cnt_CLID = 0
                    strCMD = String.Format("Select count(*) from T_SAPIM WHERE (SAPPN <> '') AND (NOT (SAPPN IS NULL))")
                    Cnt_CLID = DA.ExecuteScalar(strCMD)
                    If Not Cnt_CLID > 0 Then
                        Exit Do
                    End If
                End Using
            Loop

            ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Step 'Rollback Mapping' finish")
        Catch ex As Exception
            ErrorLogging("SAPMigration-RollbackCLID", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            RollbackCLIDInfo = False
            Exit Function
        Finally
        End Try
        RollbackCLIDInfo = True
    End Function

    Public Function batch_receipt(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
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
            p_ds.AcceptChanges()

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
                result_flag = submit_trans("account_alias_batch_receipt", TransactionID, OracleLoginData)

                If result_flag = "Y" Then

                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag)
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_trans("account_alias_batch_receipt", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString)
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString)
                'Next
                'comm.Transaction.Rollback()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds

                result_flag = del_trans("account_alias_batch_receipt", TransactionID, OracleLoginData)
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

    Public Function batch_issue(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Long, ByVal MiscType As String) As DataSet
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
            p_ds.AcceptChanges()

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
                result_flag = submit_trans("account_alias_batch_issue", TransactionID, OracleLoginData)

                If result_flag = "Y" Then

                Else
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag)
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountissue_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_trans("account_alias_batch_issue", TransactionID, OracleLoginData)
                End If
                Return p_ds
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, DR(i)("o_return_message").ToString)
                Next
                'For Each dtRow In DR
                '    ErrorLogging(MoveType, OracleLoginData.User, dtRow("o_error_mssg").ToString)
                'Next
                'comm.Transaction.Rollback()
                'comm.Connection.Close()
                'comm.Connection.Dispose()
                'comm.Dispose()
                'Return p_ds

                result_flag = del_trans("account_alias_batch_issue", TransactionID, OracleLoginData)
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

    Public Function CheckSAPPN(ByVal SAPPN As String, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result1 As Object
                result1 = da.ExecuteScalar(String.Format("select OraclePN from T_SAPPN where SAPPN = '{0}'", SAPPN))
                If result1.ToString <> "" Then
                    'If Not result1.ToString Is DBNull.Value Then
                    CheckSAPPN = result1.ToString
                Else
                    CheckSAPPN = ""
                End If
                'result2 = da.ExecuteScalar(String.Format("Select Sum(QtyBaseUOM) from T_CLMaster where Sloc = '0010' and StatusCode <> 0 and StorageType = '001' and MaterialNo = '{0}' and StorageBin NOT LIKE 'TR%'", Material))
                'GetAvlQty.WHSQty = result2.ToString
            Catch ex As Exception
                ErrorLogging("SAPMigration-CheckSAPPN", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function Check_SAPPN(ByVal SAPPN As String, ByVal OracleLoginData As ERPLogin) As SAPPN_Check
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result1, result2 As String
                result1 = da.ExecuteScalar(String.Format("select OraclePN from T_SAPPN where SAPPN = '{0}'", SAPPN))
                If Not result1 Is Nothing AndAlso Not result1 Is DBNull.Value AndAlso result1.ToString <> "" Then
                    'If Not result1.ToString Is DBNull.Value Then
                    Check_SAPPN.OraclePN = result1.ToString
                    result2 = da.ExecuteScalar(String.Format("select OraclePN from T_OraclePN where OraclePN = '{0}'", Check_SAPPN.OraclePN))
                    If Not result2 Is Nothing AndAlso Not result2 Is DBNull.Value AndAlso result2.ToString <> "" Then
                        Check_SAPPN.Flag = True
                    End If
                Else
                    Check_SAPPN.OraclePN = ""
                End If
                'result2 = da.ExecuteScalar(String.Format("Select Sum(QtyBaseUOM) from T_CLMaster where Sloc = '0010' and StatusCode <> 0 and StorageType = '001' and MaterialNo = '{0}' and StorageBin NOT LIKE 'TR%'", Material))
                'GetAvlQty.WHSQty = result2.ToString
            Catch ex As Exception
                ErrorLogging("SAPMigration-Check_SAPPN", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using

    End Function

    Public Function submit_trans(ByVal MoveType As String, ByVal TransactionID As Double, ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()

            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_mtl_tran_pkg.mtl_process_online"
                Oda.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                'Oda.Parameters.Add("p_timout", OracleType.Int32).Value = 1800000
                'Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = GetOrgCode()   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_error_mssg", OracleType.VarChar, 10000)
                Oda.Parameters("o_success_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_error_mssg").Direction = ParameterDirection.Output

                Oda.Connection.Open()
                aa = CInt(Oda.ExecuteNonQuery())
                MsgFlag = Oda.Parameters("o_success_flag").Value
                Oda.Connection.Close()
                ErrorLogging(MoveType, OracleLoginData.User, DirectCast(MsgFlag, String))
                Return DirectCast(MsgFlag, String)
            Catch oe As OracleException
                ErrorLogging(MoveType, OracleLoginData.User, "TransactionID: " & TransactionID & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function del_trans(ByVal MoveType As String, ByVal TransactionID As Double, ByVal OracleLoginData As ERPLogin) As String
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
                MsgFlag = Oda.Parameters("o_succ_flag").Value
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
End Class
