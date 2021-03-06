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


Public Class StockTake
    Inherits PublicFunction

#Region "Physical Inventory"
    Public Function Get_STCtrlList(ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using DA As DataAccess = GetDataAccess()
                Dim strCMD As String
                strCMD = String.Format("Select * from T_StockTake_Control with(NOLOCK)")
                Get_STCtrlList = DA.ExecuteDataSet(strCMD, "Control")
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-Get_STCtrlList", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function ST_CheckAction(ByVal Action As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Try
            Using DA As DataAccess = GetDataAccess()
                ST_CheckAction = DA.ExecuteScalar(String.Format("Select Enabled from T_StockTake_Control with(NOLOCK) where Action = '{0}'", Action))
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-ST_CheckAction", OracleLogindata.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function ST_CompAction(ByVal Action As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Try
            Dim strCMD, ActionList As String
            Dim rs As Integer

            Using DA As DataAccess = GetDataAccess()
                If LCase(Action) = "locketrace" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    ActionList = "CopyClmaster,CpySubLoc,ClearPIName,AddPIName,Add,CopyPI,ExtCnt2,SetCount1,EnableScan1"
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True' Where CHARINDEX(Action, '{0}') > 0", ActionList)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "copyclmaster" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "cpysubLoc" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "clearpiname" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "addpiname" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "add" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "copypi" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "extcnt2" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "setcount1" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "enablescan1" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    ActionList = "SumQty,NoValidate,Stop1Bak,WithValidate,SetCount2,EnableScan2,Stop2Bak"
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True' Where CHARINDEX(Action, '{0}') > 0", ActionList)
                    rs = DA.ExecuteNonQuery(strCMD)

                    ActionList = "ClearPIName,AddPIName,Add,CopyPI,ExtCnt2"
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where CHARINDEX(Action, '{0}') > 0", ActionList)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "sumqty" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "novalidate" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "stop1bak" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "withvalidate" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "setcount2" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "enablescan2" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "stop2bak" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    ActionList = "ChckNotFound,BkpRmv,GenCSV,CheckBlankTag,BkpBFAdjust"
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True' Where CHARINDEX(Action, '{0}') > 0", ActionList)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "chcknotfound" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "bkprmv" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "gencsv" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "checkblanktag" Then

                ElseIf LCase(Action) = "bkpbfadjust" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    ActionList = "Loss,Gain,NotFound,NewFind,DiffLocator,LossDiff,GainDiff,BkpAfter"
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True' Where CHARINDEX(Action, '{0}') > 0", ActionList)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "forloss" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "forgain" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "fornotfound" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "fornewfind" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "fordifflocator" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "forlossdiff" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "forgaindiff" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)

                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False' Where Action = '{0}'", "BkpBFAdjust")
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "bkpafteradjust" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'False', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "unlocketrace" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                ElseIf LCase(Action) = "refresh" Then
                    strCMD = String.Format("Update T_StockTake_Control Set Enabled = 'True', Done = 'True', DoneTime = getdate(), DoneBy = '{1}' Where Action = '{0}'", Action, OracleLogindata.User)
                    rs = DA.ExecuteNonQuery(strCMD)
                End If
            End Using
            Return True
        Catch ex As Exception
            ErrorLogging("StockTake-ST_CompAction", OracleLogindata.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function UpdateActionStatus(ByVal Action As String, ByVal Enabled As Boolean, ByVal Done As Boolean, ByVal OracleLogindata As ERPLogin) As Boolean
        Try
            Using DA As DataAccess = GetDataAccess()
                DA.ExecuteNonQuery(String.Format("Update T_StockTake_Control SET Enable = '{0}', Done = '{1}' where Action = '{2}'", Enabled, Done, Action))

                If Action = "LockeTrace" Then

                ElseIf Action = "CopyClmaster" Then

                ElseIf Action = "CpySubLoc" Then

                ElseIf Action = "ClearPIName" Then

                ElseIf Action = "AddPIName" Then

                ElseIf Action = "CopyPI" Then

                ElseIf Action = "ExtCnt2" Then

                ElseIf Action = "SetCount1" Then

                ElseIf Action = "EnbScan" Then



                End If
                UpdateActionStatus = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-UpdateActionStatus", oraclelogindata.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function Lock_eTrace(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-Lock_eTrace", OracleLoginData.User.ToUpper, "Step 'Lock_eTrace' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_Lock_eTrace_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            'CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-Lock_eTrace", OracleLoginData.User.ToUpper, "Step 'Lock_eTrace' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-Lock_eTrace", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        'Using da As DataAccess = GetDataAccess()
        '    Dim strCMD As String
        '    Try
        '        strCMD = String.Format("UPDATE T_Transaction SET Active = 0 where TranID not like 'INV007%'")
        '        da.ExecuteNonQuery(strCMD)

        '        strCMD = String.Format("UPDATE T_Transaction SET Active = 1 where TranID like 'INV007%'")
        '        da.ExecuteNonQuery(strCMD)

        '        Lock_eTrace = True
        '    Catch oe As Exception
        '        ErrorLogging("StockTake-Lock_eTrace", OracleLoginData.User, oe.Message & oe.Source, "E")
        '        Lock_eTrace = False
        '    End Try
        'End Using
    End Function

    Public Function UnLock_eTrace(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-UnLock_eTrace", OracleLoginData.User.ToUpper, "Step 'UnLock_eTrace' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_UnLock_eTrace_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            'CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-UnLock_eTrace", OracleLoginData.User.ToUpper, "Step 'UnLock_eTrace' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-UnLock_eTrace", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        'Using da As DataAccess = GetDataAccess()
        '    Dim strCMD As String
        '    Try
        '        strCMD = String.Format("UPDATE T_Transaction SET Active = 1 where TranID not like 'INV005%' and TranID not like 'INV007%' and TranID not like 'QAL001%' and TranID not like 'WHS005%'")
        '        da.ExecuteNonQuery(strCMD)

        '        strCMD = String.Format("UPDATE T_Transaction SET Active = 0 where TranID like 'INV005%' or TranID like 'INV007%' or TranID like 'QAL001%' or TranID like 'WHS005%'")
        '        da.ExecuteNonQuery(strCMD)

        '        UnLock_eTrace = True
        '    Catch oe As Exception
        '        ErrorLogging("StockTake-UnLock_eTrace", OracleLoginData.User, oe.Message & oe.Source, "E")
        '        UnLock_eTrace = False
        '    End Try
        'End Using
    End Function

    Public Function ClearPIName(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Try
            Using DA As DataAccess = GetDataAccess()
                DA.ExecuteNonQuery(String.Format("Delete from T_PIList"))
                ClearPIName = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-ClearPIName", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source.ToString)
            ClearPIName = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function GetOrgList_StockTake(ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim SqlStr, OrgList As String

                SqlStr = String.Format("SELECT value from T_Config where configID = 'ORGLIST'")
                OrgList = FixNull(da.ExecuteScalar(SqlStr))
                Return OrgList

            Catch ex As Exception
                ErrorLogging("StockTake-GetOrgList", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function AddPIName(ByVal OrgCode As String, ByVal PIName As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Try
            Using DA As DataAccess = GetDataAccess()
                DA.ExecuteNonQuery(String.Format("Insert into T_PIList(OrgCode, PIName, Active) Values('{0}','{1}','{2}')", OrgCode, PIName, "1"))
                AddPIName = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-AddPIName", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            AddPIName = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function StockTake_CpySubLoc(ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter             'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Try
                Dim PIList As New DataSet
                Dim PIData As New DataSet
                Dim i, j, ra As Integer
                Dim OrgList, Sqlstr As String

                Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
                Dim SQLTransaction As SqlClient.SqlTransaction
                Dim CLMasterSQLCommand As SqlClient.SqlCommand

                Sqlstr = String.Format("Delete from T_SubinvLocator")
                da.ExecuteNonQuery(Sqlstr)

                Sqlstr = String.Format("SELECT Value FROM T_Config with(nolock) WHERE  ConfigID = 'ORGLIST'")
                OrgList = FixNull(da.ExecuteScalar(Sqlstr))

                If OrgList <> "" Then
                    If InStr(OrgList, ",") > 0 Then
                        Dim Org_List() As String
                        Erase Org_List
                        Org_List = Split(OrgList, ",")
                        For i = LBound(Org_List) To UBound(Org_List)
                            Dim ds As New DataSet()
                            ds.Tables.Add("SubinvLoc")
                            oda = da.Oda_Sele()

                            Dim OrgID As String = GetOrgID(Org_List(i))

                            oda.SelectCommand.CommandType = CommandType.StoredProcedure
                            oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_org_sub_loc"

                            oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = OrgID  'Org_List(i)
                            oda.SelectCommand.Parameters.Add("o_sub_loc_data", OracleType.Cursor).Direction = ParameterDirection.Output

                            oda.SelectCommand.Connection.Open()
                            oda.Fill(ds, "SubinvLoc")
                            oda.SelectCommand.Connection.Close()

                            If Not ds Is Nothing And ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                                Try
                                    myConn.Open()
                                    SQLTransaction = myConn.BeginTransaction

                                    For j = 0 To ds.Tables(0).Rows.Count - 1
                                        Sqlstr = String.Format("INSERT into T_SubinvLocator(OrgCode,SubInv,Locator) values ('{0}','{1}','{2}')", Org_List(i), ds.Tables(0).Rows(j)("subinventory"), ds.Tables(0).Rows(j)("locator"))
                                        'da.ExecuteNonQuery(Sqlstr)
                                        CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                                        CLMasterSQLCommand.Connection = myConn
                                        CLMasterSQLCommand.Transaction = SQLTransaction
                                        ra = CLMasterSQLCommand.ExecuteNonQuery()
                                    Next
                                    SQLTransaction.Commit()
                                Catch ex As Exception
                                    StockTake_CpySubLoc = StockTake_CpySubLoc & " Insert to T_SubinvLocator with error. "
                                    SQLTransaction.Rollback()
                                    ErrorLogging("StockTake_CpySubLoc", OracleLoginData.User, ex.Message & ex.Source, "E")
                                Finally
                                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                                End Try
                            Else
                                StockTake_CpySubLoc = StockTake_CpySubLoc & " No Subinv/Locator downloaded from Org " & Org_List(i)
                                ErrorLogging("StockTake_CpySubLoc", OracleLoginData.User, "No Subinv/Locator downloaded from Org " & Org_List(i), "I")
                            End If
                        Next
                    Else
                        Dim ds As New DataSet()
                        ds.Tables.Add("SubinvLoc")
                        oda = da.Oda_Sele()

                        Dim OrgID As String = GetOrgID(OrgList)

                        oda.SelectCommand.CommandType = CommandType.StoredProcedure
                        oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_org_sub_loc"

                        oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = OrgID  'OrgList
                        oda.SelectCommand.Parameters.Add("o_sub_loc_data", OracleType.Cursor).Direction = ParameterDirection.Output

                        oda.SelectCommand.Connection.Open()
                        oda.Fill(ds, "SubinvLoc")
                        oda.SelectCommand.Connection.Close()

                        If Not ds Is Nothing And ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                            Try
                                myConn.Open()
                                SQLTransaction = myConn.BeginTransaction

                                For j = 0 To ds.Tables(0).Rows.Count - 1
                                    Sqlstr = String.Format("INSERT into T_SubinvLocator(OrgCode,SubInv,Locator) values ('{0}','{1}','{2}')", OrgList, FixNull(ds.Tables(0).Rows(j)("subinventory")), FixNull(ds.Tables(0).Rows(j)("locator")))
                                    'da.ExecuteNonQuery(Sqlstr)
                                    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                                    CLMasterSQLCommand.Connection = myConn
                                    CLMasterSQLCommand.Transaction = SQLTransaction
                                    ra = CLMasterSQLCommand.ExecuteNonQuery()
                                Next
                                SQLTransaction.Commit()
                            Catch ex As Exception
                                StockTake_CpySubLoc = StockTake_CpySubLoc & " Insert to T_SubinvLocator with error. "
                                SQLTransaction.Rollback()
                                ErrorLogging("StockTake_CpySubLoc", OracleLoginData.User, ex.Message & ex.Source, "E")
                            Finally
                                If myConn.State <> ConnectionState.Closed Then myConn.Close()
                            End Try
                        Else
                            StockTake_CpySubLoc = StockTake_CpySubLoc & " No Subinv/Locator downloaded from Org " & OrgList
                            ErrorLogging("StockTake_CpySubLoc", OracleLoginData.User, "No Subinv/Locator downloaded from Org " & OrgList, "I")
                        End If
                    End If

                Else
                    StockTake_CpySubLoc = StockTake_CpySubLoc & " Setup error for OrgList in T_Config"
                    ErrorLogging("StockTake_CpySubLoc", OracleLoginData.User, "Setup error for OrgList in T_Config", "I")
                End If
            Catch oe As Exception
                StockTake_CpySubLoc = StockTake_CpySubLoc & oe.Message & oe.Source
                ErrorLogging("Get_SubInv_Restrict", "", oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
            Return StockTake_CpySubLoc
        End Using
    End Function

    Public Function CopyPI(ByVal OracleLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter             'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num
            Try
                Dim PIList As New DataSet
                Dim PIData As New DataSet
                Dim i, j, ra As Integer
                Dim Sqlstr As String

                Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
                Dim SQLTransaction As SqlClient.SqlTransaction
                Dim CLMasterSQLCommand As SqlClient.SqlCommand

                Sqlstr = String.Format("DELETE from T_PIData")
                da.ExecuteNonQuery(Sqlstr)

                Sqlstr = String.Format("SELECT OrgCode, PIName  FROM T_PIList WHERE  Active = 'True'")
                PIList = da.ExecuteDataSet(Sqlstr, "PIList")

                If Not PIList Is Nothing And PIList.Tables.Count > 0 And PIList.Tables(0).Rows.Count > 0 Then
                    For i = 0 To PIList.Tables(0).Rows.Count - 1
                        Dim ds As New DataSet()
                        oda = da.Oda_Sele()
                        oda.SelectCommand.CommandType = CommandType.StoredProcedure
                        oda.SelectCommand.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_pi_tag_data"              'Get Standard PO

                        oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = PIList.Tables(0).Rows(i)("OrgCode")
                        oda.SelectCommand.Parameters.Add("p_pi_name", OracleType.VarChar, 50).Value = PIList.Tables(0).Rows(i)("PIName")
                        oda.SelectCommand.Parameters.Add("o_pitag_data", OracleType.Cursor).Direction = ParameterDirection.Output

                        oda.SelectCommand.Connection.Open()
                        oda.Fill(ds)
                        oda.SelectCommand.Connection.Close()

                        If Not ds Is Nothing And ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                            Try
                                myConn.Open()
                                SQLTransaction = myConn.BeginTransaction

                                For j = 0 To ds.Tables(0).Rows.Count - 1
                                    If FixNull(ds.Tables(0).Rows(j)("item_num")) <> "" Then
                                        Sqlstr = String.Format("INSERT into T_PIData(BuCode,OrgCode,PIName,TagType,TagNo,Item,Rev,MaterialDesc,SubInv,Locator,Lot,ExpDate,UOM,Cost,BookQty) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')", ds.Tables(0).Rows(j)("bu_code"), ds.Tables(0).Rows(j)("org_code"), ds.Tables(0).Rows(j)("pi_name"), "NON-BLANK", ds.Tables(0).Rows(j)("tag_number"), ds.Tables(0).Rows(j)("item_num"), ds.Tables(0).Rows(j)("revision"), SQLString(ds.Tables(0).Rows(j)("item_description")), ds.Tables(0).Rows(j)("subinventory"), ds.Tables(0).Rows(j)("locator"), ds.Tables(0).Rows(j)("lot_number"), ds.Tables(0).Rows(j)("lot_expiration_date"), ds.Tables(0).Rows(j)("uom"), ds.Tables(0).Rows(j)("item_cost"), ds.Tables(0).Rows(j)("system_quantity"))
                                        'da.ExecuteNonQuery(Sqlstr)
                                        CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                                        CLMasterSQLCommand.Connection = myConn
                                        CLMasterSQLCommand.Transaction = SQLTransaction
                                        ra = CLMasterSQLCommand.ExecuteNonQuery()
                                    End If
                                Next
                                SQLTransaction.Commit()
                            Catch ex As Exception
                                CopyPI = CopyPI & " Insert to T_PIData with error. "
                                SQLTransaction.Rollback()
                                ErrorLogging("StockTake-CopyPI", OracleLoginData.User, ex.Message & ex.Source, "E")
                            Finally
                                If myConn.State <> ConnectionState.Closed Then myConn.Close()
                            End Try
                        Else
                            CopyPI = CopyPI & " No record for " & "PI: " & PIList.Tables(0).Rows(i)("PIName")
                            ErrorLogging("StockTake-CopyPI", OracleLoginData.User, "No record download for PI: " & PIList.Tables(0).Rows(i)("PIName"), "I")
                        End If
                    Next
                Else
                    CopyPI = CopyPI & " No PI to be copied. "
                    ErrorLogging("StockTake-CopyPI", OracleLoginData.User, "No PI is available to copy record", "I")
                End If
            Catch oe As Exception
                CopyPI = CopyPI & oe.Message & oe.Source
                ErrorLogging("StockTake-CopyPI", "", oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
            Return CopyPI
        End Using
    End Function

    Public Function CopyCLMaster(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-CopyCLMaster", OracleLoginData.User.ToUpper, "Step 'CopyCLMaster' start", "I")

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
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_copy_clmaster_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            'CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-CopyCLMaster", OracleLoginData.User.ToUpper, "Step 'CopyCLMaster' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-CopyCLMaster", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ExtCnt2(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim Adpt As SqlDataAdapter
        Dim Sqlstr As String
        Dim flag, ra As Integer

        Dim dsSumSubInv, dsSumItem As DataSet
        Dim i, j As Integer
        Dim sumqty, perct As Decimal
        Dim subinv As String

        'ErrorLogging("StockTake-CopyCLMaster", OracleLoginData.User.ToUpper, "Step 'ExtCnt2' start", "I")

        'Using da As DataAccess = GetDataAccess()
        '    Try
        '        Sqlstr = String.Format("delete from T_StockTake_Cnt2")
        '        da.ExecuteNonQuery(sqlstr)

        '        sqlstr = String.Format("select OrgCode, PIName, SubInv, SUM(BookQty * Cost) as 'SumSubInv' from T_PIData where TagType = 'NON-BLANK' Group By OrgCode, PIName, SubInv")
        '        dsSumSubInv = New DataSet
        '        dsSumSubInv = da.ExecuteDataSet(sqlstr, "SumBySubInv")

        '        For i = 0 To dsSumSubInv.Tables(0).Rows.Count - 1
        '            sqlstr = String.Format("select OrgCode, PIName, SubInv, Item, SUM(BookQty * Cost) as SumItem from T_PIData where TagType = 'NON-BLANK' and OrgCode = '{0}' and PIName = '{1}' and SubInv = '{2}'  Group By OrgCode, PIName, SubInv, Item Order By SumItem DESC", dsSumSubInv.Tables(0).Rows(i)("OrgCode"), dsSumSubInv.Tables(0).Rows(i)("PIName"), dsSumSubInv.Tables(0).Rows(i)("SubInv"))
        '            dsSumItem = New DataSet
        '            dsSumItem = da.ExecuteDataSet(sqlstr, "SumByItem")

        '            For j = 0 To dsSumItem.Tables(0).Rows.Count - 1
        '                sumqty = sumqty + dsSumItem.Tables(0).Rows(j)("SumItem")
        '                perct = sumqty / dsSumSubInv.Tables(0).Rows(i)("SumSubInv")
        '                'dsSumItem.Tables(0).Rows(j)("Cnt2Flag") = "Y"
        '                'dsSumItem.Tables(0).AcceptChanges()
        '                sqlstr = String.Format("Select Subinventory from T_SubInvList where Type = '{0}' and Subinventory = '{1}'", "BF", dsSumSubInv.Tables(0).Rows(i)("SubInv"))
        '                subinv = da.ExecuteScalar(sqlstr)
        '                If subinv Is Nothing OrElse subinv = "" Then
        '                    sqlstr = String.Format("Insert into T_StockTake_Cnt2(OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot, CLID, StatusCode, BoxID) select T_PIData.OrgCode, T_PIData.PIName, T_PIData.TagType, T_PIData.TagNo, T_PIData.SubInv, T_PIData.Locator, T_PIData.Item, T_PIData.Rev, T_PIData.Lot, T_StockTake.CLID, T_StockTake.StatusCode, T_StockTake.BoxID from T_PIData left outer join T_StockTake on T_PIData.OrgCode = T_StockTake.OrgCode and T_PIData.Subinv = T_StockTake.Subinv and T_PIData.Locator = T_StockTake.Locator and T_PIData.Item = T_StockTake.MaterialNo and T_PIData.Rev = T_StockTake.MaterialRevision and T_PIData.Lot = T_StockTake.RTLot where T_PIData.TagType = 'NON-BLANK' and T_PIData.OrgCode = '{0}' and T_PIData.PIName = '{1}' and T_PIData.SubInv = '{2}' and T_PIData.Item = '{3}' ", dsSumItem.Tables(0).Rows(j)("OrgCode"), dsSumItem.Tables(0).Rows(j)("PIName"), dsSumItem.Tables(0).Rows(j)("SubInv"), dsSumItem.Tables(0).Rows(j)("Item"))
        '                    da.ExecuteNonQuery(sqlstr)
        ''                Else
        ''                    sqlstr = String.Format("Insert into T_StockTake_Cnt2(OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot) select OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot from T_PIData where T_PIData.TagType = 'NON-BLANK' and T_PIData.OrgCode = '{0}' and T_PIData.PIName = '{1}' and T_PIData.SubInv = '{2}' and T_PIData.Item = '{3}'", dsSumItem.Tables(0).Rows(j)("OrgCode"), dsSumItem.Tables(0).Rows(j)("PIName"), dsSumItem.Tables(0).Rows(j)("SubInv"), dsSumItem.Tables(0).Rows(j)("Item"))
        ''                    da.ExecuteNonQuery(sqlstr)
        '                End If

        '                If perct >= 0.3 Then
        '                    Exit For
        '                End If
        '            Next
        '        Next
        '        sqlstr = String.Format("Delete From T_StockTake where CLID <> NULL and StatusCode <> '1'")
        '        da.ExecuteNonQuery(sqlstr)

        '        Return True
        '    Catch ex As Exception
        '        Throw ex
        '        ErrorLogging("StockTake-ExtCnt2", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        '        Return False
        '    End Try
        'End Using

        'Step 1.
        ErrorLogging("StockTake-ExtCnt2", OracleLoginData.User.ToUpper, "Step 'ExtCnt2'' start", "I")
        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            Sqlstr = String.Format("delete from T_StockTake_Cnt2")
            'da.ExecuteNonQuery(Sqlstr)
            CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            Sqlstr = String.Format("select OrgCode, PIName, SubInv, SUM(BookQty * Cost) as 'SumSubInv' from T_PIData where TagType = 'NON-BLANK' Group By OrgCode, PIName, SubInv")
            dsSumSubInv = New DataSet
            'da.ExecuteNonQuery(Sqlstr)
            Adpt = New SqlDataAdapter(Sqlstr, myConn)
            Adpt.SelectCommand.Connection = myConn
            Adpt.SelectCommand.Transaction = SQLTransaction
            'Adpt.UpdateCommand.Connection = myConn
            'Adpt.UpdateCommand.Transaction = SQLTransaction
            Adpt.Fill(dsSumSubInv, "SumBySubInv")

            For i = 0 To dsSumSubInv.Tables(0).Rows.Count - 1
                Sqlstr = String.Format("select OrgCode, PIName, SubInv, Item, SUM(BookQty * Cost) as SumItem from T_PIData where TagType = 'NON-BLANK' and OrgCode = '{0}' and PIName = '{1}' and SubInv = '{2}'  Group By OrgCode, PIName, SubInv, Item Order By SumItem DESC", dsSumSubInv.Tables(0).Rows(i)("OrgCode"), dsSumSubInv.Tables(0).Rows(i)("PIName"), dsSumSubInv.Tables(0).Rows(i)("SubInv"))
                dsSumItem = New DataSet
                'da.ExecuteNonQuery(Sqlstr)
                Adpt = New SqlDataAdapter(Sqlstr, myConn)
                Adpt.SelectCommand.Connection = myConn
                Adpt.SelectCommand.Transaction = SQLTransaction
                'Adpt.UpdateCommand.Connection = myConn
                'Adpt.UpdateCommand.Transaction = SQLTransaction
                Adpt.Fill(dsSumItem, "SumByItem")

                sumqty = 0
                perct = 0
                For j = 0 To dsSumItem.Tables(0).Rows.Count - 1
                    sumqty = sumqty + dsSumItem.Tables(0).Rows(j)("SumItem")
                    perct = CDec(sumqty / dsSumSubInv.Tables(0).Rows(i)("SumSubInv"))
                    'dsSumItem.Tables(0).Rows(j)("Cnt2Flag") = "Y"
                    'dsSumItem.Tables(0).AcceptChanges()
                    Sqlstr = String.Format("Select Subinventory from T_SubInvList where Type = '{0}' and Subinventory = '{1}'", "BF", dsSumSubInv.Tables(0).Rows(i)("SubInv"))
                    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                    CLMasterSQLCommand.Connection = myConn
                    CLMasterSQLCommand.Transaction = SQLTransaction
                    subinv = CLMasterSQLCommand.ExecuteScalar()

                    If subinv Is Nothing OrElse subinv = "" Then
                        Sqlstr = String.Format("Insert into T_StockTake_Cnt2(OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot, CLID, StatusCode,BoxID) select T_PIData.OrgCode, T_PIData.PIName, T_PIData.TagType, T_PIData.TagNo, T_PIData.SubInv, T_PIData.Locator, T_PIData.Item, T_PIData.Rev, T_PIData.Lot, T_StockTake.CLID, T_StockTake.StatusCode, T_StockTake.BoxID from T_PIData left outer join T_StockTake on T_PIData.OrgCode = T_StockTake.OrgCode and T_PIData.Subinv = T_StockTake.Subinv and T_PIData.Locator = T_StockTake.Locator and T_PIData.Item = T_StockTake.MaterialNo and T_PIData.Rev = T_StockTake.MaterialRevision and T_PIData.Lot = T_StockTake.RTLot where T_PIData.TagType = 'NON-BLANK' and T_PIData.OrgCode = '{0}' and T_PIData.PIName = '{1}' and T_PIData.SubInv = '{2}' and T_PIData.Item = '{3}' ", dsSumItem.Tables(0).Rows(j)("OrgCode"), dsSumItem.Tables(0).Rows(j)("PIName"), dsSumItem.Tables(0).Rows(j)("SubInv"), dsSumItem.Tables(0).Rows(j)("Item"))
                        CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                        CLMasterSQLCommand.Connection = myConn
                        CLMasterSQLCommand.Transaction = SQLTransaction
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        'Else
                        '    Sqlstr = String.Format("Insert into T_StockTake_Cnt2(OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot) select OrgCode, PIName, TagType, TagNo, SubInv, Locator, Item, Rev, Lot from T_PIData where T_PIData.TagType = 'NON-BLANK' and T_PIData.OrgCode = '{0}' and T_PIData.PIName = '{1}' and T_PIData.SubInv = '{2}' and T_PIData.Item = '{3}'", dsSumItem.Tables(0).Rows(j)("OrgCode"), dsSumItem.Tables(0).Rows(j)("PIName"), dsSumItem.Tables(0).Rows(j)("SubInv"), dsSumItem.Tables(0).Rows(j)("Item"))
                        '    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
                        '    CLMasterSQLCommand.Connection = myConn
                        '    CLMasterSQLCommand.Transaction = SQLTransaction
                        '    ra = CLMasterSQLCommand.ExecuteNonQuery()
                    End If

                    If perct >= 0.3 Then
                        Exit For
                    End If
                Next

            Next
            Sqlstr = String.Format("Delete From T_StockTake_Cnt2 where CLID IS NOT NULL AND CLID <> '' and StatusCode <> '1'")
            CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
            ErrorLogging("StockTake-ExtCnt2", OracleLoginData.User.ToUpper, "Step 'ExtCnt2'' finish", "I")
            Return True
        Catch ex As Exception
            flag = -1
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ExtCnt2", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function SetCount1(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Try
            Using DA As DataAccess = GetDataAccess()
                DA.ExecuteNonQuery(String.Format("Update T_Config SET Value = 'Count 1' where ConfigID = 'MMT006'"))
                SetCount1 = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-SetCount1", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            SetCount1 = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function SetCount2(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Try
            Using DA As DataAccess = GetDataAccess()
                DA.ExecuteNonQuery(String.Format("Update T_Config SET Value = 'Count 2' where ConfigID = 'MMT006'"))
                SetCount2 = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-SetCount2", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            SetCount2 = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function CheckCountOption(ByVal OracleLogindata As ERPLogin) As String
        Try
            Using DA As DataAccess = GetDataAccess()
                CheckCountOption = Convert.ToString(DA.ExecuteScalar(String.Format("select Value from T_Config with(nolock) where ConfigID = 'MMT006'")))
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-CheckCountOption", OracleLogindata.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            CheckCountOption = ex.Message.ToString & ex.Source.ToString
            Exit Function
        Finally
        End Try
    End Function

    Public Function StockTake_ValidateSubLoc(ByVal OracleLoginData As ERPLogin, ByVal Subinv As String, ByVal Locator As String) As String
        Dim i, j As Integer
        Dim rs, strCMD As String
        Try
            Using DA As DataAccess = GetDataAccess()
                If Locator = "" Then
                    strCMD = String.Format("Select count(Subinv) as Cnt from T_SubinvLocator with(nolock) where OrgCode = '{0}' and Subinv = '{1}'", OracleLoginData.OrgCode, Subinv)
                    i = FixNull(DA.ExecuteScalar(strCMD))
                    If i = 0 Then
                        StockTake_ValidateSubLoc = "Invalid Subinventory: " & Subinv & "!"
                    ElseIf i > 0 Then
                        strCMD = String.Format("Select count(*) as Cnt from T_SubinvLocator with(nolock) where OrgCode = '{0}' and Subinv = '{1}' and Locator <> ''", OracleLoginData.OrgCode, Subinv)
                        j = FixNull(DA.ExecuteScalar(strCMD))
                        If j > 0 Then
                            StockTake_ValidateSubLoc = "Y"
                        ElseIf j = 0 Then
                            StockTake_ValidateSubLoc = "N"
                        End If
                    End If
                ElseIf Locator <> "" Then
                    strCMD = String.Format("Select count(Locator) as Cnt from T_SubinvLocator with(nolock) where OrgCode = '{0}' and Subinv = '{1}' and Locator = '{2}'", OracleLoginData.OrgCode, Subinv, Locator)
                    i = FixNull(DA.ExecuteScalar(strCMD))
                    If i > 0 Then
                        StockTake_ValidateSubLoc = "Y"
                    Else
                        StockTake_ValidateSubLoc = "Invalid Locator: " & Locator & "!"
                    End If
                End If
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-StopScan", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            StockTake_ValidateSubLoc = "Invalid Subinv/Locator"
            Exit Function
        Finally
        End Try
    End Function

    Public Function CheckInSubLocList(ByVal LoginData As ERPLogin, ByVal SubInv As String, ByVal Locator As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim result As String
            Try
                result = Convert.ToString(da.ExecuteScalar(String.Format("select Subinventory from T_SubInvList with(nolock) where Subinventory = '{0}' and PI_Flag = '{1}'", SubInv, True)))
                If Not result Is Nothing And result <> "" Then
                    CheckInSubLocList = "Y"
                Else
                    CheckInSubLocList = "N"
                End If
            Catch oe As Exception
                ErrorLogging("StockTake-CheckInSubLocList", LoginData.User, "SubInventory: " & SubInv & ", " & oe.Message & oe.Source, "E")
                CheckInSubLocList = "N"
            End Try
            Return CheckInSubLocList
        End Using
    End Function

    Public Function CheckBFSubinv(ByVal Subinv As String, ByVal OracleLogindata As ERPLogin) As Boolean
        Dim strcmd As String
        Dim result As Object
        Try
            Using da As DataAccess = GetDataAccess()
                result = da.ExecuteScalar(String.Format("SELECT Type from T_SubinvList with(nolock) Where SubInventory = '{0}'", Subinv.ToUpper))
                If FixNull(result.ToString) = "BF" Then
                    CheckBFSubinv = True
                Else
                    CheckBFSubinv = False
                End If
            End Using
        Catch ex As Exception
            CheckBFSubinv = False
            ErrorLogging("StockTake-CheckBFSubinv", OracleLogindata.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function GetCLIDInfo(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim ds As DataSet
                Dim Sqlstr As String

                ds = New DataSet
                Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,SubInv,Locator,RTLot,StatusCode,SupplyType,BoxID,C1_By,C1_Qty,C1_SubInv,C1_Locator,C2_By,C2_Qty,C2_SubInv,C2_Locator from T_StockTake with(nolock) where CLID = '{0}' and OrgCode = '{1}'", CLID, OracleLoginData.OrgCode)
                ds = da.ExecuteDataSet(Sqlstr, "CLID")

                If ds Is Nothing OrElse ds.Tables.Count < 1 OrElse ds.Tables(0).Rows.Count < 1 Then
                    Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo,MaterialRevision,Qty,UOM,SubInv,Locator,RTLot,StatusCode,SupplyType,BoxID,C1_By,C1_Qty,C1_SubInv,C1_Locator,C2_By,C2_Qty,C2_SubInv,C2_Locator from T_StockTake with(nolock) where BoxID = '{0}' and OrgCode = '{1}' and StatusCode = '1'", CLID, OracleLoginData.OrgCode)
                    ds = da.ExecuteDataSet(Sqlstr, "CLID")
                End If

                Return ds
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-GetCLIDInfo", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function CheckStopFlag(ByVal OracleLoginData As ERPLogin) As String
        Try
            Using da As DataAccess = GetDataAccess()
                CheckStopFlag = Convert.ToString(da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT003'")))
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-CheckStopFlag", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            CheckStopFlag = ""
            Exit Function
        End Try
    End Function

    Public Function Save_STChange(ByVal pstInfo As DataSet, ByVal CountSeq As String, ByVal OracleLoginData As ERPLogin) As String
        Dim i As Integer
        Dim strCMD, Stop_Scan, Remarks As String
        Try
            Using DA As DataAccess = GetDataAccess()

                Stop_Scan = Convert.ToString(DA.ExecuteScalar(String.Format("select Value from T_Config with(nolock) where ConfigID = 'MMT003'")))
                If Stop_Scan = "YES" Then
                    Save_STChange = "Have stopped scanning!"
                    Exit Function
                End If

                Remarks = OracleLoginData.User.ToUpper & " remove counting history at " & Date.Now.ToString
                ErrorLogging("StockTake-Save_STChange1", OracleLoginData.User.ToUpper, "Remarks: " & Remarks, "I")

                If CountSeq = "Count 1" Then
                    For i = 0 To pstInfo.Tables("CLID").Rows.Count - 1
                        If pstInfo.Tables("CLID").Rows(i)("Variance") = "RemoveHistory" Then
                            ErrorLogging("StockTake-Save_STChange2", OracleLoginData.User.ToUpper, "Remarks: " & Remarks, "I")

                            strCMD = String.Format("UPDATE T_StockTake with(rowlock) set C1_By = NULL, C2_By = NULL, C1_On = NULL, C2_On = NULL, C1_Qty = NULL, C2_Qty = NULL, C_Qty = NULL , C1_SubInv = NULL, C2_SubInv = NULL, C_SubInv = NULL, C1_Locator = NULL, C2_Locator = NULL, C_Locator = NULL, C1_StatusCode = NULL, C2_StatusCode = NULL, C_StatusCode = NULL, C_SupplyType = NULL, Variance = NULL, Remarks = '{1}'   WHERE (CLID = '{0}') ", pstInfo.Tables("CLID").Rows(i)("CLID"), Remarks)
                            DA.ExecuteNonQuery(strCMD)
                        Else
                            strCMD = String.Format("UPDATE T_StockTake with(rowlock) set C1_By = '{0}', C1_On = getdate(),C1_Qty = '{1}', C_Qty = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN c_qty ELSE  '{1}' END) ,C1_SubInv = '{2}', C_SubInv = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN C_SubInv ELSE  '{2}' END),C1_Locator = '{3}',C_Locator = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN C_Locator ELSE  '{3}' END),C1_StatusCode = '{4}',C_StatusCode = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN C_StatusCode ELSE  '{4}' END),C_SupplyType = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN C_SupplyType ELSE  '{5}' END),Variance = (CASE WHEN c2_by <> '' AND NOT c2_by IS NULL THEN Variance ELSE  '{6}' END)   WHERE (CLID = '{7}') ", OracleLoginData.User, pstInfo.Tables("CLID").Rows(i)("C_Qty"), pstInfo.Tables("CLID").Rows(i)("C_SubInv"), pstInfo.Tables("CLID").Rows(i)("C_Locator"), pstInfo.Tables("CLID").Rows(i)("C_StatusCode"), pstInfo.Tables("CLID").Rows(i)("C_SupplyType"), pstInfo.Tables("CLID").Rows(i)("Variance"), pstInfo.Tables("CLID").Rows(i)("CLID"))
                            DA.ExecuteNonQuery(strCMD)
                        End If
                    Next
                ElseIf CountSeq = "Count 2" Then
                    For i = 0 To pstInfo.Tables("CLID").Rows.Count - 1
                        If pstInfo.Tables("CLID").Rows(i)("Variance") = "RemoveHistory" Then
                            ErrorLogging("StockTake-Save_STChange3", OracleLoginData.User.ToUpper, "Remarks: " & Remarks, "I")

                            strCMD = String.Format("UPDATE T_StockTake with(rowlock) set C1_By = NULL, C2_By = NULL, C1_On = NULL, C2_On = NULL, C1_Qty = NULL, C2_Qty = NULL, C_Qty = NULL , C1_SubInv = NULL, C2_SubInv = NULL, C_SubInv = NULL, C1_Locator = NULL, C2_Locator = NULL, C_Locator = NULL, C1_StatusCode = NULL, C2_StatusCode = NULL, C_StatusCode = NULL, C_SupplyType = NULL, Variance = NULL, Remarks = '{1}'   WHERE (CLID = '{0}') ", pstInfo.Tables("CLID").Rows(i)("CLID"), Remarks)
                            DA.ExecuteNonQuery(strCMD)
                        Else
                            strCMD = String.Format("UPDATE T_StockTake with(rowlock) set C2_By = '{0}', C2_On = getdate(),C2_Qty = '{1}', C_Qty = '{1}',C2_SubInv = '{2}', C_SubInv = '{2}', C2_Locator = '{3}',C_Locator = '{3}',C2_StatusCode = '{4}',C_StatusCode = '{4}',C_SupplyType = '{5}',Variance = '{6}'   WHERE (CLID = '{7}') ", OracleLoginData.User, pstInfo.Tables("CLID").Rows(i)("C_Qty"), pstInfo.Tables("CLID").Rows(i)("C_SubInv"), pstInfo.Tables("CLID").Rows(i)("C_Locator"), pstInfo.Tables("CLID").Rows(i)("C_StatusCode"), pstInfo.Tables("CLID").Rows(i)("C_SupplyType"), pstInfo.Tables("CLID").Rows(i)("Variance"), pstInfo.Tables("CLID").Rows(i)("CLID"))
                            DA.ExecuteNonQuery(strCMD)
                        End If
                    Next
                End If
                Save_STChange = ""
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-Save_STChange", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            Save_STChange = ex.Message.ToString & ex.Source.ToString
            Exit Function
        Finally
        End Try
    End Function

    Public Function CheckNotFound(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim strCMD As String


        ErrorLogging("StockTake-CheckNotFound", OracleLoginData.User.ToUpper, "Step 'CheckNotFound' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            'strCMD = String.Format("UPDATE T_STOCKTAKE set Variance = 'Not Found' WHERE StatusCode = '1' and (C1_By = '' or C1_By IS NULL) and (SubInv IN (SELECT Subinventory from T_SubInvList where Type <> 'BF' and PI_Flag = 'True') )")
            strCMD = String.Format("UPDATE T_STOCKTAKE set Variance = 'Not Found' WHERE StatusCode = '1' and (Variance is NULL or Variance = '') and (SubInv IN (SELECT Subinventory from T_SubInvList with(nolock) where Type <> 'BF' and PI_Flag = 'True') )")

            CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            SQLTransaction.Commit()
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-CheckNotFound", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            CheckNotFound = False
            Exit Function
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        CheckNotFound = True
        ErrorLogging("StockTake-CheckNotFound", OracleLoginData.User.ToUpper, "Step 'CheckNotFound' finish", "I")
    End Function

    Public Function GenVarRpt(ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Sqlstr = String.Format("Select * from T_StockTake with(nolock) where SubInv = 'ZR0 SM1' or SubInv = 'ZR0 FS1' and Locator in ('SIK1KB....','SIK2KB....','SIS2KB....')")
                Return da.ExecuteDataSet(Sqlstr, "CLID")
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-GenVarRpt", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function EnableScan(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Dim strCMD As String
        Try
            Using DA As DataAccess = GetDataAccess()
                strCMD = String.Format("UPDATE T_Config set Value = 'NO' WHERE (ConfigID = 'MMT003')")
                DA.ExecuteNonQuery(strCMD)
                EnableScan = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-EnableScan", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            EnableScan = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function StopScan(ByVal SeqNo As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Dim rs, strCMD As String
        Try
            Using DA As DataAccess = GetDataAccess()
                strCMD = String.Format("sp_stop_and_backup_stocktake '{0}'", SeqNo)
                rs = DA.ExecuteScalar(strCMD)

                If rs = "0" Then
                    StopScan = False
                ElseIf rs = "1" Then
                    StopScan = True
                End If
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-StopScan" & "-" & SeqNo, OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            StopScan = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function NoValidate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Dim strCMD As String
        Try
            Using DA As DataAccess = GetDataAccess()
                strCMD = String.Format("UPDATE T_Config set Value = '2' WHERE (ConfigID = 'SYS001')")
                DA.ExecuteNonQuery(strCMD)
                NoValidate = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-NoValidate", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            NoValidate = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function WithValidate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i As Integer
        Dim strCMD As String
        Try
            Using DA As DataAccess = GetDataAccess()
                strCMD = String.Format("UPDATE T_Config set Value = '1' WHERE (ConfigID = 'SYS001')")
                DA.ExecuteNonQuery(strCMD)
                WithValidate = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-WithValidate", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            WithValidate = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function BkpRmv(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-BkpRmv", OracleLoginData.User.ToUpper, "Step 'Backup and Remove Unconcerned Records' start", "I")

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
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_backup_remove_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-BkpRmv", OracleLoginData.User.ToUpper, "Step 'Backup and Remove Unconcerned Records' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-BkpRmv", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function SumQtyForPI(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_sumqty_to_pitable", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        'Dim SQLTransaction As SqlClient.SqlTransaction
        'Dim CLMasterSQLCommand As SqlClient.SqlCommand
        'Dim Sqlstr As String
        'Dim flag, ra As Integer

        ''Step 1.
        'ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Initial Table' start", "I")
        'Try
        '    myConn.Open()
        '    SQLTransaction = myConn.BeginTransaction
        '    Sqlstr = String.Format("DELETE from T_StockTake_Sum")
        '    'da.ExecuteNonQuery(Sqlstr)
        '    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
        '    CLMasterSQLCommand.Connection = myConn
        '    CLMasterSQLCommand.Transaction = SQLTransaction
        '    ra = CLMasterSQLCommand.ExecuteNonQuery()
        '    SQLTransaction.Commit()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Initial Table' finish", "I")
        'Catch ex As Exception
        '    flag = -1
        '    SQLTransaction.Rollback()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        '    Return False
        'Finally
        '    If myConn.State <> ConnectionState.Closed Then myConn.Close()
        'End Try

        ''Step 2.
        'ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Sum CountedQty' start", "I")
        'Try
        '    myConn.Open()
        '    SQLTransaction = myConn.BeginTransaction
        '    'Sqlstr = String.Format("INSERT into T_StockTake_Sum(OrgCode,MaterialNo,MaterialRevision,SubInv,Locator,RTLot,CntQty) select OrgCode,MaterialNo,MaterialRevision,SubInv,Locator,RTLot, SUM(CountedQty) from T_StockTake where Variance <> '' and not Variance is NULL Group By OrgCode,MaterialNo,MaterialRevision,SubInv,Locator,RTLot")
        '    Sqlstr = String.Format("INSERT into T_StockTake_Sum select OrgCode,MaterialNo,MaterialRevision,SubInv,Locator,RTLot, SUM(CountedQty) as CntQty from T_StockTake where Variance <> '' and not Variance is NULL Group By OrgCode,MaterialNo,MaterialRevision,SubInv,Locator,RTLot")
        '    'da.ExecuteNonQuery(Sqlstr)
        '    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
        '    CLMasterSQLCommand.Connection = myConn
        '    CLMasterSQLCommand.Transaction = SQLTransaction
        '    ra = CLMasterSQLCommand.ExecuteNonQuery()
        '    SQLTransaction.Commit()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Sum CountedQty' finish", "I")
        'Catch ex As Exception
        '    flag = -1
        '    SQLTransaction.Rollback()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        '    Return False
        'Finally
        '    If myConn.State <> ConnectionState.Closed Then myConn.Close()
        'End Try

        ''Step 3.
        'ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Update Qty to PI Table' start", "I")
        'Try
        '    myConn.Open()
        '    SQLTransaction = myConn.BeginTransaction
        '    Sqlstr = String.Format("UPDATE T_PIData SET CountedQty = T_StockTake_SUM.CntQty FROM T_StockTake_SUM INNER JOIN T_PIData ON T_StockTake_SUM.OrgCode = T_PIData.OrgCode AND T_StockTake_SUM.MaterialNo = T_PIData.Item AND T_StockTake_SUM.MaterialRevision = T_PIData.Rev AND T_StockTake_SUM.SubInv = T_PIData.SubInv AND T_StockTake_SUM.Locator = T_PIData.Locator AND T_StockTake_SUM.RTLot = T_PIData.Lot")
        '    'da.ExecuteNonQuery(Sqlstr)
        '    CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
        '    CLMasterSQLCommand.Connection = myConn
        '    CLMasterSQLCommand.Transaction = SQLTransaction
        '    ra = CLMasterSQLCommand.ExecuteNonQuery()
        '    SQLTransaction.Commit()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, "Step 'Sum Counted Qty to PI Table - Update Qty to PI Table' finish", "I")
        'Catch ex As Exception
        '    flag = -1
        '    SQLTransaction.Rollback()
        '    ErrorLogging("StockTake-SumQtyForPI", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        '    Return False
        'Finally
        '    If myConn.State <> ConnectionState.Closed Then myConn.Close()
        'End Try
    End Function

    Public Function StockTake_AssignExpDate(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim i, j As Integer
        Dim strCMD As String

        Try
            Using DA As DataAccess = GetDataAccess()
                Dim ds As New DataSet
                Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
                Dim comm As OracleCommand = DA.Ora_Command_Trans()
                Dim aa As OracleString
                Dim resp As Integer = OracleLoginData.RespID_Inv   '53485     'CAROLD3    53330
                Dim appl As Integer = OracleLoginData.AppID_Inv    '401

                strCMD = String.Format("Select OrgCode as p_org_code, item as p_item_num, lot as p_lot_num, NULL as o_lot_control_code, NULL as o_shelf_life_ctrl, '' as o_exp_date from T_PIData with(nolock) where TagType = 'BLANK'")
                ds = DA.ExecuteDataSet(strCMD, "BlankTag_List")

                For Each drow As DataRow In ds.Tables(0).Rows
                    drow.SetAdded()
                Next

                If Not ds Is Nothing And ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                    comm.CommandType = CommandType.StoredProcedure
                    comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                    comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(OracleLoginData.UserID)  ''15784
                    comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = resp
                    comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = appl
                    comm.ExecuteOracleNonQuery(aa)
                    comm.Parameters.Clear()

                    comm.CommandText = "apps.xxetr_inv_mtl_tran_pkg.get_lot_expiration_date"   '53330   401
                    comm.Parameters.Add("p_org_code", OracleType.VarChar, 240)
                    comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
                    comm.Parameters.Add("p_lot_num", OracleType.VarChar, 240)
                    comm.Parameters.Add("o_lot_control_code", OracleType.Number)
                    comm.Parameters.Add("o_shelf_life_ctrl", OracleType.Number)
                    comm.Parameters.Add("o_exp_date", OracleType.VarChar, 240)

                    comm.Parameters("o_lot_control_code").Direction = ParameterDirection.Output
                    comm.Parameters("o_shelf_life_ctrl").Direction = ParameterDirection.Output
                    comm.Parameters("o_exp_date").Direction = ParameterDirection.Output

                    comm.Parameters("p_org_code").SourceColumn = "p_org_code"
                    comm.Parameters("p_item_num").SourceColumn = "p_item_num"
                    comm.Parameters("p_lot_num").SourceColumn = "p_lot_num"
                    comm.Parameters("o_lot_control_code").SourceColumn = "o_lot_control_code"
                    comm.Parameters("o_shelf_life_ctrl").SourceColumn = "o_shelf_life_ctrl"
                    comm.Parameters("o_exp_date").SourceColumn = "o_exp_date"

                    oda_h.InsertCommand = comm
                    oda_h.Update(ds.Tables("BlankTag_List"))

                    For i = 0 To ds.Tables("BlankTag_List").Rows.Count - 1
                        strCMD = String.Format("UPDATE T_PIData SET ExpDate = '{0}' where OrgCode = '{1}' and Item = '{2}' and Lot = '{3}'", ds.Tables("BlankTag_List").Rows(i)("o_exp_date"), ds.Tables("BlankTag_List").Rows(i)("p_org_code"), ds.Tables("BlankTag_List").Rows(i)("p_item_num"), ds.Tables("BlankTag_List").Rows(i)("p_lot_num"))
                        DA.ExecuteNonQuery(strCMD)
                    Next
                End If
                StockTake_AssignExpDate = True
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake_AssignExpDate", OracleLoginData.User.ToUpper, "Error: " & ex.Message & ex.Source, "E")
            StockTake_AssignExpDate = False
            Exit Function
        Finally
        End Try
    End Function

    Public Function UpdateDiffLocator(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-UpdateDiffLocator", OracleLoginData.User.ToUpper, "Step 'Update for Diff Locator Records' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_Update_DiffLocator", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-UpdateDiffLocator", OracleLoginData.User.ToUpper, "Step 'Update for Diff Locator Records' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-UpdateDiffLocator", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetPIResult(ByVal OracleLoginData As ERPLogin) As DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim ra As String
        Dim SqlStr As String

        ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, "Step 'Generate CSV file - Set ZERO for no-onhand PI record' start", "I")
        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction
            SqlStr = String.Format("UPDATE T_PIData SET CountedQty = 0 where CountedQty is NULL")
            'da.ExecuteNonQuery(Sqlstr)
            CLMasterSQLCommand = New SqlClient.SqlCommand(Sqlstr, myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            ra = CLMasterSQLCommand.ExecuteNonQuery()
            SQLTransaction.Commit()
        Catch ex As Exception
            ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        End Try
        ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, "Step 'Generate CSV file - Set ZERO for no-onhand PI record' finish", "I")



        ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, "Step 'Generate CSV file - Collect Data' start", "I")
        Try
            Using DA As DataAccess = GetDataAccess()
                Dim sqlcmd As String
                sqlcmd = String.Format("Select BuCode as Bu_Code,OrgCode as Org_Code,PIName as PI_Name,TagType as Tag_Type,TagNo as Tag_Number,Item, Rev as Revision, SubInv as Subinventory, Locator, Lot as Lot_Number, UOM, ExpDate as Expires_On, CountedQty as Quantity from T_PIData with(nolock) where Item IS NOT NULL and Item <> '' order by PIName, TagNo")
                Return DA.ExecuteDataSet(sqlcmd, "transfer_table")
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
        End Try
        ErrorLogging("StockTake-GetPIResult", OracleLoginData.User.ToUpper, "Step 'Generate CSV file - Collect Data' finish", "I")
    End Function

    Public Function StockTake_BkpBFAdjust(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake_BkpBFAdjust", OracleLoginData.User.ToUpper, "Step 'Copy T_CLMaster to T_CLMaster_Before_Adjustment' start", "I")

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
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_backup_clmaster_before_adjust_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake_BkpBFAdjust", OracleLoginData.User.ToUpper, "Step 'Copy T_CLMaster to T_CLMaster_Before_Adjustment' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake_BkpBFAdjust", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ForLoss(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForLoss", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Loss Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForLoss", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForLoss", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Loss Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForLoss", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
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

    Public Function SubInvTransfer(ByVal p_ds As DataSet, ByVal OracleLoginData As ERPLogin, ByVal MoveType As String, ByVal TransactionID As Integer) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim oda_h As OracleDataAdapter = New OracleDataAdapter()
        Dim comm As OracleCommand = da.Ora_Command_Trans()
        Dim aa As OracleString
        Dim resp As Integer = OracleLoginData.RespID_Inv
        Dim appl As Integer = OracleLoginData.AppID_Inv

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
            'comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
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
            'comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
            comm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
            comm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
            comm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
            comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
            comm.Parameters("p_transaction_type_name").SourceColumn = "p_transaction_type_name"
            comm.Parameters("o_return_status").SourceColumn = "o_return_status"
            comm.Parameters("o_return_message").SourceColumn = "o_return_message"

            oda_h.InsertCommand = comm


            'ErrorLogging("StockTake_SubInvTransfer-", OracleLoginData.User, DStoXML(p_ds), "I")
            oda_h.Update(p_ds.Tables("transfer_table"))
            Dim DR() As DataRow = Nothing
            Dim i As Integer
            DR = p_ds.Tables("transfer_table").Select("o_return_status = 'N'")

            If DR.Length = 0 Then
                Dim PostTransfer As String
                'PostTransfer = Submit_Transfer(OracleLoginData, TransactionID, MoveType)
                If PostTransfer <> "Y" Then
                    If PostTransfer = "" Then PostTransfer = "Post failed"
                    p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("transfer_table").Rows(0)("o_return_message") = PostTransfer    '"Post failed"
                    ErrorLogging("StockTake_SubInvTransfer-" & MoveType, OracleLoginData.User, PostTransfer, "I")
                    del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
                End If
            Else
                For i = 0 To (DR.Length - 1)
                    ErrorLogging("StockTake_SubInvTransfer-" & MoveType, OracleLoginData.User, DR(i)("o_return_message").ToString, "I")
                Next
                del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
            End If

            comm.Connection.Close()
            Return p_ds

        Catch ex As Exception
            p_ds.Tables("transfer_table").Rows(0)("o_return_status") = "N"
            p_ds.Tables("transfer_table").Rows(0)("o_return_message") = ex.Message
            ErrorLogging("StockTake_SubInvTransfer-" & MoveType, OracleLoginData.User, ex.Message & ex.Source, "E")
            del_transfer_inte(CInt(OracleLoginData.UserID), TransactionID)
            Return p_ds
        Finally
            If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
        End Try
    End Function

    Public Function ForGain(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForGain", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Gain Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForGain", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForGain", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Gain Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForGain", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
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

            '' No need p_lot_expiration_date
            ''comm.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
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
            ''comm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
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
                    ErrorLogging(MoveType & "-Post", OracleLoginData.User, result_flag, "I")
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_status") = "N"
                    p_ds.Tables("accountreceipt_table").Rows(0)("o_return_message") = result_flag
                    result_flag = del_trans("account_alias_batch_receipt", TransactionID, OracleLoginData)
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
                ErrorLogging(MoveType, OracleLoginData.User, DirectCast(MsgFlag, String), "I")
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

    Public Function ForNotFound(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForNotFound", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Not Found Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForNotFound", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForNotFound", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Not Found Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForNotFound", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ForNewFind(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForNewFind", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for New Find Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForNewFind", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForNewFind", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for New Find Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForNewFind", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ForDiffLocator(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForDiffLocator", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Diff Locator Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForDiffLocator", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForDiffLocator", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Diff Locator Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForDiffLocator", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ForLossDiff(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForLossDiff", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Loss & Diff Locator Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForLossDiff", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForLossDiff", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Loss & Diff Locator Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForLossDiff", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ForGainDiff(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake-ForGainDiff", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Gain & Diff Locator Record' start", "I")

        Try
            myConn.Open()
            SQLTransaction = myConn.BeginTransaction

            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_StockTake_ForGainDiff", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@User", SqlDbType.VarChar, 50)).Value = "StockTake"
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake-ForGainDiff", OracleLoginData.User.ToUpper, "Step 'Update T_CLMaster for Gain & Diff Locator Record' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake-ForGainDiff", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function StockTake_BkpAfterAdjust(ByVal OracleLoginData As ERPLogin) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SQLTransaction As SqlClient.SqlTransaction
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim rs As String
        Dim strCMD As String

        ErrorLogging("StockTake_BkpAfterAdjust", OracleLoginData.User.ToUpper, "Step 'Copy T_CLMaster to T_CLMaster_After_Adjustment' start", "I")

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
            CLMasterSQLCommand = New SqlClient.SqlCommand("sp_backup_clmaster_after_adjust_for_stocktake", myConn)
            CLMasterSQLCommand.Connection = myConn
            CLMasterSQLCommand.CommandType = CommandType.StoredProcedure
            CLMasterSQLCommand.Parameters.Add(New SqlParameter("@OrgCode", SqlDbType.VarChar, 50)).Value = OracleLoginData.OrgCode
            CLMasterSQLCommand.Transaction = SQLTransaction
            CLMasterSQLCommand.CommandTimeout = 1000 * 60 * 240
            rs = CLMasterSQLCommand.ExecuteScalar.ToString
            SQLTransaction.Commit()

            ErrorLogging("StockTake_BkpAfterAdjust", OracleLoginData.User.ToUpper, "Step 'Copy T_CLMaster to T_CLMaster_Before_Adjustment' finish", "I")

            If rs = "0" Then
                Return False
            ElseIf rs = "1" Then
                Return True
            End If
        Catch ex As Exception
            SQLTransaction.Rollback()
            ErrorLogging("StockTake_BkpAfterAdjust", OracleLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function
#End Region

#Region "BF Count"

    Public Function BFC_GetWeekCode(ByVal OracleLoginData As ERPLogin) As String
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr, Year, Week As String
                Sqlstr = String.Format("Select datepart(yyyy, getdate()-1) as Year")
                Year = da.ExecuteScalar(Sqlstr)
                Sqlstr = String.Format("Select datepart(wk, getdate()-1) as Week")
                Week = da.ExecuteScalar(Sqlstr)
                BFC_GetWeekCode = Year & "-" & Week
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-GetWeekCode", OracleLoginData.User, ex.Message & ex.Source, "E")
            BFC_GetWeekCode = ""
        End Try
        Return BFC_GetWeekCode
    End Function

    Public Function BFC_GetBFSubinv(ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Sqlstr = String.Format("Select ProcessID, Description from T_SysLov where Name = 'BF Count' ")
                BFC_GetBFSubinv = da.ExecuteDataSet(Sqlstr, "BFSubinv")
            End Using
        Catch ex As Exception
            ErrorLogging("StockTake-GetBFSubinv", OracleLoginData.User, ex.Message & ex.Source, "E")
            BFC_GetBFSubinv = Nothing
        End Try
    End Function

    Public Function BFC_DelOldBFCount(ByVal WeekCode As String, ByVal BFSubinv As String, ByVal OracleLoginData As ERPLogin) As Boolean
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Sqlstr = String.Format("Delete from T_BFCount where WeekCode <> '{0}' and OrgCode = '{1}' and Subinv = '{2}'", WeekCode, OracleLoginData.OrgCode, BFSubinv)
                da.ExecuteNonQuery(Sqlstr)
                Return True
            End Using
        Catch ex As Exception
            ErrorLogging("BFC-GetBFSubinv", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

    Public Function BFC_GetCLIDInfo(ByVal CLID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo as Item,QtyBaseUOM as Qty, SLOC as SubInv,RTLot, 'Add' as Remarks from T_CLMaster with(nolock) where CLID = '{0}' and OrgCode = '{1}'", CLID, OracleLoginData.OrgCode)
                Return da.ExecuteDataSet(Sqlstr, "CLID")
            End Using
        Catch ex As Exception
            ErrorLogging("BFC_GetCLIDInfo", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return Nothing
        End Try
    End Function

    Public Function Save_BFCount(ByVal WeekCode As String, ByVal lblInfo As DataSet, ByVal OracleLoginData As ERPLogin) As Boolean
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr As String
                Dim rs As Object
                If lblInfo.Tables(0).Rows(0)("Remarks").ToString = "Add" Then
                    Sqlstr = String.Format("Select CLID from T_BFCount where WeekCode = '{0}' and OrgCode = '{1}' and CLID = '{2}'", WeekCode, lblInfo.Tables(0).Rows(0)("OrgCode"), lblInfo.Tables(0).Rows(0)("CLID"))
                    rs = da.ExecuteScalar(Sqlstr)
                    If rs Is Nothing Then
                        Sqlstr = String.Format("Insert into T_BFCount values('{0}','{1}','{2}','{3}','{4}','{5}',getdate(),'{6}','')", WeekCode, lblInfo.Tables(0).Rows(0)("OrgCode"), lblInfo.Tables(0).Rows(0)("Subinv"), lblInfo.Tables(0).Rows(0)("CLID"), lblInfo.Tables(0).Rows(0)("Item"), lblInfo.Tables(0).Rows(0)("Qty"), OracleLoginData.User)
                    Else
                        Sqlstr = String.Format("Update T_BFCount SET Subinv = '{0}', Item = '{1}', Qty = '{2}', ChangedOn=getdate(),ChangedBy = '{3}' where WeekCode = '{4}' and OrgCode = '{5}' and CLID = '{6}'", lblInfo.Tables(0).Rows(0)("Subinv"), lblInfo.Tables(0).Rows(0)("Item"), lblInfo.Tables(0).Rows(0)("Qty"), OracleLoginData.User, WeekCode, lblInfo.Tables(0).Rows(0)("OrgCode"), lblInfo.Tables(0).Rows(0)("CLID"))
                    End If
                ElseIf lblInfo.Tables(0).Rows(0)("Remarks").ToString = "Del" Then
                    Sqlstr = String.Format("delete from T_BFCount where weekcode = '{0}' and OrgCode = '{1}' and CLID = '{2}'", WeekCode, lblInfo.Tables(0).Rows(0)("OrgCode"), lblInfo.Tables(0).Rows(0)("CLID"))
                End If
                da.ExecuteNonQuery(Sqlstr)
                Return True
            End Using
        Catch ex As Exception
            ErrorLogging("BFC_GetCLIDInfo", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return False
        End Try
    End Function

#End Region
End Class
