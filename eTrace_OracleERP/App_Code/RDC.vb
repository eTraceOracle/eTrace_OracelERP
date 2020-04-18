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




Public Class RDC
    Inherits PublicFunction

#Region "RDC"

    Public Function ReadFlow(ByVal IntSN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("select rtrim(b.Process) as 'Process',rtrim(b.FailedUnitReturnTo) as 'FailedUnitReturnTo' from T_WIPHeader a,T_WIPFlow b where a.WIPID=b.WIPID and a.IntSN='{0}' order by b.SeqNo", IntSN)
                Return da.ExecuteDataSet(strSQL, "ProcessList")
            Catch ex As Exception
                ErrorLogging("RDC-ReadFlow", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadFailData(ByVal IntSN) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RepGetFailData '{0}'", IntSN)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("RDC-ReadFailData", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadRepairCode() As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_ReadRepCodes")
                Return da.ExecuteDataSet(strSQL, "RepairCode")
            Catch ex As Exception
                ErrorLogging("RDC-ReadRepairCode", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadRepCodesByCategory(ByVal Category As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_ReadRepCodesByCategory '{0}'", Category)
                Return da.ExecuteDataSet(strSQL, "RepairCode")
            Catch ex As Exception
                ErrorLogging("RDC-ReadRepCodesByCategory", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadRepDefectCodeByCause(ByVal Cause As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_ReadRepDefectCodeByCause '{0}'", Cause)
                Return da.ExecuteDataSet(strSQL, "RepairCode")
            Catch ex As Exception
                ErrorLogging("RDC-ReadRepDefectCodeByCause", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadRepairData(ByVal IntSN As String) As DataSet
        Dim dsTmp As New DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RepGetRepairData '{0}'", IntSN)
                dsTmp = da.ExecuteDataSet(strSQL, 2)
                dsTmp.Tables(0).TableName = "RepDataItem"
                dsTmp.Tables(1).TableName = "RepDataItemDetail"
                Return dsTmp
            Catch ex As Exception
                ErrorLogging("RDC-ReadRepairData", "", ex.Message)
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function ReadNDFData(ByVal IntSN As String, ByVal OperatorName As String, ByVal Type As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                If Type.ToUpper = "FailData".ToUpper Then
                    strSQL = String.Format("exec sp_RepGetNDFData '" + IntSN + "','','','FailData','',''")
                    Return da.ExecuteDataSet(strSQL)
                Else
                    Dim NewRepID As String
                    Dim OldRepID As String
                    Dim RepTime As String
                    Dim dsTmp As DataSet
                    Dim dsData As New DataSet
                    Dim dtHeader As Data.DataTable
                    Dim dtItem As Data.DataTable
                    Dim dtDetail As Data.DataTable
                    strSQL = String.Format("exec sp_RepGetNDFData '" + IntSN + "','','','NewID','',''")
                    dsTmp = da.ExecuteDataSet(strSQL)
                    If dsTmp Is Nothing Then
                        Return Nothing
                    End If
                    If dsTmp.Tables Is Nothing Then
                        Return Nothing
                    End If
                    If dsTmp.Tables(0).Rows.Count < 1 Then
                        Return Nothing
                    End If
                    NewRepID = dsTmp.Tables(0).Rows(0)(0).ToString
                    OldRepID = dsTmp.Tables(0).Rows(0)(1).ToString
                    RepTime = StandardTime()
                    strSQL = String.Format("exec sp_RepGetNDFData '" + IntSN + "','" + NewRepID + "','','Header','',''")
                    dtHeader = da.ExecuteDataSet(strSQL, "RepHeader").Tables("RepHeader").Copy
                    dsData.Tables.Add(dtHeader)
                    strSQL = String.Format("exec sp_RepGetNDFData '" + IntSN + "','" + NewRepID + "','" + OldRepID + "','Item',N'" + OperatorName + "','" + RepTime + "'")
                    dtItem = da.ExecuteDataSet(strSQL, "RepDataItem").Tables("RepDataItem").Copy
                    dsData.Tables.Add(dtItem)
                    strSQL = String.Format("exec sp_RepGetNDFData '" + IntSN + "','" + NewRepID + "','" + OldRepID + "','Detail','',''")
                    dtDetail = da.ExecuteDataSet(strSQL, "RepDataItemDetail").Tables("RepDataItemDetail").Copy
                    dsData.Tables.Add(dtDetail)
                    Return dsData
                End If
            Catch ex As Exception
                ErrorLogging("RDC-ReadNDFData", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function GetCLIDData(ByVal CLID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim strSQL As String
                strSQL = String.Format("select a.CLID,a.MaterialNo,a.DateCode,a.LotNo,a.VendorName,a.Manufacturer, a.StatusCode, b.PO from T_CLMaster a left join T_PO_CLID b on a.CLID=b.CLID where a.CLID='{0}'", CLID)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("RDC-GetCLIDData", "", ex.Message & ex.Source, "E")
                Return New Dataset
            End Try
        End Using
    End Function

    Public Function ReadWIPFlow(ByVal IntSN As String) As DataSet
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_ReadWIPFlow '{0}'", IntSN)
            Return da.ExecuteDataSet(strSQL)
        Catch ex As Exception
            ErrorLogging("RDC-ReadWIPFlow", "", ex.Message & ex.Source, "E")
        End Try

    End Function

    Public Function SaveRepairRecordData(ByVal RepData As DataSet, ByVal OTO As Boolean, ByVal IntSN As String) As Boolean
        Dim da As DataAccess = GetDataAccess()
        Dim xmlData As String
        xmlData = RepData.GetXml()
        xmlData = xmlData.Replace("<BurnTime>0</BurnTime>", "<BurnTime>NULL</BurnTime>")
        xmlData = xmlData.Replace("+00:00", "")
        xmlData = xmlData.Replace("+08:00", "")
        xmlData = xmlData.Replace(".0000000", "")
        Try
            Dim strSQL As String
            Dim dsSave As DataSet
            strSQL = String.Format("exec sp_RepSaveRepairRecord N'{0}','{1}','{2}'", xmlData, OTO, IntSN)
            dsSave = da.ExecuteDataSet(strSQL)
            If dsSave Is Nothing Then
                Return False
            End If
            If dsSave.Tables Is Nothing Then
                Return False
            End If
            If dsSave.Tables(0).Rows.Count < 1 Then
                Return False
            End If
            If dsSave.Tables(0).Rows(0).Item(0).ToString.Trim = "0" Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            ErrorLogging("RDC-SaveRepairDetailData", "", ex.Message & ex.Source, "E")
            Return False
        End Try

    End Function

    Public Function RepRaiseTimes(ByVal ds As DataSet) As String
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Try
            strSQL = String.Format("exec sp_RepRaiseTimes N'{0}'", DStoXML(ds))
            Return da.ExecuteScalar(strSQL).ToString
        Catch ex As Exception
            ErrorLogging("RDC-RepRaiseTimes", "", ex.Message & ex.Source, "E")
            Return strSQL '"Save Fail!"
        End Try


    End Function
    Public Function ServerTime() As DateTime
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim TimefromDB As Date
        Try
            strSQL = String.Format("select getdate()")
            ServerTime = Convert.ToDateTime(da.ExecuteScalar(strSQL))
            'StandardTime = Format(TimefromDB, "MM-dd-yyyy HH:mm:ss") & "." & Format(TimefromDB.Millisecond, "000")
        Catch ex As Exception
            'StandardTime = Format(Now(), "MM-dd-yyyy HH:mm:ss")
            ErrorLogging("SFC-StadardTime", "", ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function CheckOTO(ByVal SN As String, ByVal Type As String) As Boolean
        CheckOTO = False
        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim sResult As String
        Dim objReader As SqlClient.SqlDataReader
        Try
            strSQL = String.Format("exec SP_RepCheckOTO '{0}','{1}'", SN, Type)
            objReader = da.ExecuteReader(strSQL)
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then sResult = objReader.GetValue(0)
            End While
            If sResult = "0" Then
                CheckOTO = True
            ElseIf sResult = "1" Then
                CheckOTO = False
            End If
        Catch ex As Exception
            ErrorLogging("RDC-CheckOTO", "", ex.Message)
            Return False
        End Try
    End Function

    Public Function ReadStructure(ByVal Model As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("select PCBA from T_ProductStructure where Model='{0}'", Model)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("RDC-ReadStructure", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using
    End Function

    Public Function ReadConfig(ByVal ConfigID As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim dsReturn As New DataSet
            Dim strSQL As String
            Try
                strSQL = String.Format("select ConfigID from T_Config where Value='YES' and ConfigID='{0}'", ConfigID)
                dsReturn = da.ExecuteDataSet(strSQL)
                If dsReturn Is Nothing Then
                    Return False
                Else
                    If dsReturn.Tables Is Nothing Then
                        Return False
                    Else
                        If dsReturn.Tables(0).Rows.Count > 0 Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
            Catch ex As Exception
                ErrorLogging("RDC-ReadConfig", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function SaveRepairDetailData(ByVal RepData As DataSet, ByVal OTO As Boolean, ByVal IntSN As String) As Boolean

        Dim sRepId As String
        Dim iItem As Integer
        Dim iTmpItem As Integer
        Dim strSQL As String
        If RepData.Tables("RepHeader").Rows(0).Item("RepID").ToString.Trim = "" Then
            Using da As DataAccess = GetDataAccess()

                Try
                    strSQL = String.Format("Select newid() as NextID")
                    sRepId = Convert.ToString(da.ExecuteScalar(strSQL))
                Catch ex As Exception
                    ErrorLogging("RDC-SaveRepairDetailData", "", ex.Message & ex.Source, "E")
                    Return False
                End Try
            End Using
        Else
            sRepId = RepData.Tables("RepHeader").Rows(0).Item("RepID").ToString.Trim
        End If
        RepData.Tables("RepHeader").Rows(0).Item("RepID") = sRepId

        iTmpItem = 0
        For iItem = 0 To RepData.Tables("RepDataItem").Rows.Count - 1
            If RepData.Tables("RepDataItem").Rows(iItem).Item("RepID").ToString.Trim = "" And (RepData.Tables("RepDataItem").Rows(iItem).Item("Item").ToString.Trim = "") Then
                RepData.Tables("RepDataItem").Rows(iItem).Item("RepID") = sRepId
                RepData.Tables("RepDataItem").Rows(iItem).Item("Item") = "1"
                iTmpItem = 1
                Exit For
            ElseIf RepData.Tables("RepDataItem").Rows(iItem).Item("RepID").ToString.Trim <> "" And (RepData.Tables("RepDataItem").Rows(iItem).Item("Item").ToString.Trim = "") Then
                RepData.Tables("RepDataItem").Rows(iItem).Item("Item") = (iTmpItem + 1).ToString
                iTmpItem = iTmpItem + 1
                Exit For
            Else
                iTmpItem = RepData.Tables("RepDataItem").Rows(iItem).Item("Item").ToString.Trim
            End If
        Next
        'iTmpRep = 0
        If Not (RepData.Tables("RepDataItemDetail") Is Nothing) Then
            Dim iTmpAll, iTmpDetail, iTmpDItem As Integer
            iTmpAll = 0
            iTmpDItem = 0
            iTmpDetail = 0
            Dim sItem As String = ""
            For iItem = 0 To RepData.Tables("RepDataItemDetail").Rows.Count - 1
                If RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepID").ToString.Trim = "" And (RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim = "") And RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem").ToString.Trim = "" Then
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepID") = sRepId
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item") = "1"
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem") = iTmpAll + 1
                    iTmpAll = iTmpAll + 1
                ElseIf RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepID").ToString.Trim <> "" And (RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim = "") And RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem").ToString.Trim = "" Then
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item") = (iTmpItem + 1).ToString
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem") = iTmpDItem + 1
                    iTmpDItem = iTmpDItem + 1
                ElseIf RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepID").ToString.Trim <> "" And RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim <> "" And (RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem").ToString.Trim = "") Then
                    If sItem <> RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim Then
                        sItem = RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim
                        iTmpDetail = 0
                    End If
                    RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem") = iTmpDetail + 1
                    iTmpDetail = iTmpDetail + 1
                Else
                    If sItem <> RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim Then
                        sItem = RepData.Tables("RepDataItemDetail").Rows(iItem).Item("Item").ToString.Trim
                    End If
                    iTmpDetail = RepData.Tables("RepDataItemDetail").Rows(iItem).Item("RepItem").ToString.Trim
                End If
            Next
        End If
        RepData.AcceptChanges()
        Dim xmlData As String
        xmlData = RepData.GetXml()
        xmlData = xmlData.Replace("<BurnTime>0</BurnTime>", "<BurnTime>NULL</BurnTime>")
        xmlData = xmlData.Replace("+00:00", "")
        Using da As DataAccess = GetDataAccess()
            Dim dsSave As DataSet
            Try
                strSQL = String.Format("exec sp_RepSaveRepairData N'{0}','{1}','{2}'", xmlData, OTO, IntSN)
                dsSave = da.ExecuteDataSet(strSQL)
                If dsSave Is Nothing Then
                    Return False
                End If
                If dsSave.Tables Is Nothing Then
                    Return False
                End If
                If dsSave.Tables(0).Rows.Count < 1 Then
                    Return False
                End If
                If dsSave.Tables(0).Rows(0).Item(0).ToString.Trim = "0" Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                ErrorLogging("RDC-SaveRepairDetailData", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using

        If OTO Then
            Try
                Dim iMax As Integer = 0
                Dim dsTmp As New DataSet
                For iItem = 0 To RepData.Tables("RepDataItem").Rows.Count - 1
                    If iMax <= Int(RepData.Tables("RepDataItem").Rows(iItem)("Item").ToString) Then
                        iMax = Int(RepData.Tables("RepDataItem").Rows(iItem)("Item").ToString)
                    End If
                Next
                For iItem = 0 To RepData.Tables("RepDataItem").Rows.Count - 1
                    If iMax = Int(RepData.Tables("RepDataItem").Rows(iItem)("Item").ToString) Then
                        If RepData.Tables("RepDataItem").Rows(iItem)("Status").ToString.ToUpper = "Replaced".ToUpper Then
                            dsTmp.Tables.Add(RepData.Tables("RepHeader").Copy)
                            Dim dtTmpItem As New DataTable
                            Dim tdTmpDetail As New DataTable
                            dtTmpItem.TableName = "RepDataItem"
                            dtTmpItem.Rows.Add(RepData.Tables("RepDataItem").Select("Item='" + iMax.ToString + "'"))
                            tdTmpDetail.TableName = "RepDataItemDetail"
                            tdTmpDetail.Rows.Add(RepData.Tables("RepDataItemDetail").Select("Item='" + iMax.ToString + "'"))
                            dsTmp.Tables.Add(dtTmpItem)
                            dsTmp.Tables.Add(tdTmpDetail)

                            xmlData = RepData.GetXml()
                            xmlData = xmlData.Replace("<BurnTime>0</BurnTime>", "<BurnTime>NULL</BurnTime>")
                            xmlData = xmlData.Replace("+00:00", "")
                            strSQL = "exec sp_SaveRepData N'" + xmlData + "','" + IntSN + "'"
                            Dim cmdSQL As SqlClient.SqlCommand
                            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceAdditionConnString"))
                            myConn.Open()
                            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                            cmdSQL.ExecuteScalar.ToString()
                        End If
                    End If
                Next
            Catch ex As Exception

            End Try
        End If
    End Function

    Public Function StandardTime() As String
        'Try
        '    StandardTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
        'Catch ex As Exception
        '    StandardTime = Format(DateAdd(DateInterval.Hour, 8, Now()), "MM-dd-yyyy HH:mm:ss")
        '    ErrorLogging("SFC-StadardTime", "", ex.Message & ex.Source, "E")
        'Finally

        'End Try

        Dim da As DataAccess = GetDataAccess()
        Dim strSQL As String
        Dim TimefromDB As Date
        Try
            strSQL = String.Format("select getdate()")
            TimefromDB = Convert.ToDateTime(da.ExecuteScalar(strSQL))
            StandardTime = Format(TimefromDB, "MM-dd-yyyy HH:mm:ss") & "." & Format(TimefromDB.Millisecond, "000")
        Catch ex As Exception
            'StandardTime = Format(Now(), "MM-dd-yyyy HH:mm:ss")
            ErrorLogging("SFC-StadardTime", "", ex.Message)
        End Try
    End Function

    Public Function RepScrap(ByVal IntSN As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RepScrap '{0}'", IntSN)
                Return True
            Catch ex As Exception
                ErrorLogging("RDC-RepScrap", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function SkipBI(ByVal IntSN As String) As String
        Dim da As DataAccess = GetDataAccess()
        Try
            Return da.ExecuteScalar(String.Format("exec sp_RepSkipBurnIn '{0}'", IntSN)).ToString
        Catch ex As Exception
            ErrorLogging("RDC-SkipBI", "", ex.Message)
            Return 0
        End Try
    End Function

    Public Function ReadFailItem(ByVal IntSN As String) As DataSet
        Dim dsTmp As New DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim dtHeader As Data.DataTable
            Dim dtItem As Data.DataTable
            Try
                strSQL = String.Format("select '' as ExtSerial,a.IntSerialNo as IntSerial,a.ProcessName as Process,a.SeqNo,a.PO,a.Model,a.PCBA,a.ProdDate,a.Result,a.OperatorName,a.TesterNo as Tester,a.ProgramName,a.ProgramRevision as ProgramVersion,a.IPSNo,a.IPSRevision as IPSVersion,a.Remark from T_WIPTDHeader a,T_WIPHeader b where b.IntSN='{0}' and a.WIPID=b.WIPID", IntSN)
                dtHeader = da.ExecuteDataSet(strSQL, "ResultHeader").Tables("ResultHeader").Copy
                dsTmp.Tables.Add(dtHeader)
                strSQL = String.Format("select '' as ExtSerialNo,a.IntSerialNo as IntSerialNo,a.ProcessName,a.SeqNo,b.TestStep,b.TestName,b.OutputName,b.InputCondition,b.OutputLoading,b.LowerLimit,b.Result, b.UpperLimit, b.Unit, b.Status, b.IPSReference, b.TestID from T_WIPTDHeader a,T_WIPTDItem b,T_WIPHeader c where a.TDID=b.TDID and c.WIPID=a.WIPID and c.IntSN='{0}'", IntSN)
                dtItem = da.ExecuteDataSet(strSQL, "ResultItem").Tables("ResultItem").Copy
                dsTmp.Tables.Add(dtItem)
                Return dsTmp
            Catch ex As Exception
                ErrorLogging("RDC-ReadFailItem", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function FailRecord(ByVal SN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCFailRecord '{0}'", SN)
                Return da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("RDC-FailRecord", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function NewFailData(ByVal SN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCNewFailData '{0}'", SN)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("RDC-NewFailData", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function RDCWIPFLow(ByVal WIPID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCWIPFlow '{0}'", WIPID)
                Return da.ExecuteDataSet(strSQL)
            Catch ex As Exception
                ErrorLogging("RDC-RDCWIPFLow", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function RDCSave(ByVal DS As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Dim xmlData As String
            xmlData = DS.GetXml()
            'xmlData = xmlData.Replace("<BurnTime>0</BurnTime>", "<BurnTime>NULL</BurnTime>")
            xmlData = xmlData.Replace("+00:00", "")
            xmlData = xmlData.Replace("+08:00", "")
            xmlData = xmlData.Replace(".0000000", "")
            xmlData = xmlData.Replace("'", "''")
            Try
                strSQL = String.Format("exec sp_RDCSave N'{0}'", xmlData)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("RDC-RDCSave", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function

    Public Function GetSpecifySeatItem(ByVal PCB As String, ByVal RefD As String, ByVal DJ As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim ds As New DataSet
            Dim Items As DataTable
            Items = New DataTable("Items")

            'myDataColumn = New Data.DataColumn("ASSEMBLY_ITEM", System.Type.GetType("System.String"))
            'Items.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("COMPONENT_ITEM", System.Type.GetType("System.String"))
            'Items.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("DESCRIPTION", System.Type.GetType("System.String"))
            'Items.Columns.Add(myDataColumn)

            ds.Tables.Add(Items)

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_bom_pkg.get_specify_seat_item_v2"
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = "580"
                oda.SelectCommand.Parameters.Add("p_model", OracleType.VarChar, 1000).Value = PCB
                oda.SelectCommand.Parameters.Add("p_seat", OracleType.VarChar, 1000).Value = RefD
                oda.SelectCommand.Parameters.Add("p_dj_num", OracleType.VarChar, 1000).Value = DJ

                oda.SelectCommand.Parameters.Add("o_item_bom", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 1000).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "Items")
                oda.SelectCommand.Connection.Close()
                GetSpecifySeatItem = ds
            Catch oe As Exception
                ErrorLogging("TDCService-GetSpecifySeatItem", "", "PCB:RefD " & PCB & ":" & RefD & "," & oe.Message & oe.Source)
                GetSpecifySeatItem = ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetSpecifySeatItemByOrg(ByVal PCB As String, ByVal RefD As String, ByVal DJ As String, ByVal Org As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim ds As New DataSet
            Dim Items As DataTable
            Items = New DataTable("Items")

            ds.Tables.Add(Items)

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_bom_pkg.get_specify_seat_item_v2"
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = Org
                oda.SelectCommand.Parameters.Add("p_model", OracleType.VarChar, 1000).Value = PCB
                oda.SelectCommand.Parameters.Add("p_seat", OracleType.VarChar, 1000).Value = RefD
                oda.SelectCommand.Parameters.Add("p_dj_num", OracleType.VarChar, 1000).Value = DJ

                oda.SelectCommand.Parameters.Add("o_item_bom", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_err_msg", OracleType.VarChar, 1000).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "Items")
                oda.SelectCommand.Connection.Close()
                GetSpecifySeatItemByOrg = ds
            Catch oe As Exception
                ErrorLogging("TDCService-GetSpecifySeatItem", "", "PCB:RefD:Org " & PCB & ":" & RefD & "," & Org & oe.Message & oe.Source)
                GetSpecifySeatItemByOrg = ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function RDCScrap(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCScrap N'{0}'", IntSN)
                Return da.ExecuteScalar(strSQL).ToString()
            Catch ex As Exception
                ErrorLogging("RDC-RDCScrap", "", ex.Message & ex.Source, "E")
                Return "^TDC-61@"
            End Try
        End Using
    End Function

    Public Function RDCSaveII(ByVal DS As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim result As String = ""
            Dim xmlData As String
            xmlData = DS.GetXml()
            'xmlData = xmlData.Replace("<BurnTime>0</BurnTime>", "<BurnTime>NULL</BurnTime>")
            xmlData = xmlData.Replace("+00:00", "")
            xmlData = xmlData.Replace("+08:00", "")
            xmlData = xmlData.Replace(".0000000", "")
            xmlData = xmlData.Replace("'", "''")
            Try
                strSQL = String.Format("exec sp_RDCSaveII N'{0}'", xmlData)
                result = da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("RDC-RDCSaveII", "", ex.Message & ex.Source, "E")
            End Try
            Return result
        End Using
    End Function



    Public Function FailRecordII(ByVal SN As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCFailRecordII '{0}'", SN)
                Return da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("RDC-FailRecord", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

    Public Function RDCScrapII(ByVal IntSN As String, ByVal ChangedBy As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCScrapII '{0}','{1}'", IntSN, ChangedBy)
                Return da.ExecuteScalar(strSQL).ToString
            Catch ex As Exception
                ErrorLogging("RDC-RDCScrapII", "", ex.Message & ex.Source, "E")
                Return "^TDC-61@"
            End Try
        End Using
    End Function

    Public Function RDC121MatlInfo(ByVal PCBID As String, ByVal Model As String, ByVal PCBA As String, ByVal Circuit As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDC121MatlInfo '{0}','{1}','{2}','{3}'", PCBID, Model, PCBA, Circuit)
                Return da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("RDC-RDC121MatlInfo", "", ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function


    Public Function RDCMatInfo(ByVal DJ As String, ByVal MaterialNo As String, ByVal PCBA As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_RDCMatInfo '{0}','{1}','{2}'", DJ, MaterialNo, PCBA)
                Return da.ExecuteDataSet(strSQL, 2)
            Catch ex As Exception
                ErrorLogging("RDC-RDCMatInfo", DJ + MaterialNo, ex.Message & ex.Source, "E")
                Return New DataSet
            End Try
        End Using
    End Function

#End Region

    ''' <summary>
    ''' Get Production Control info by DJ
    ''' add by michael
    ''' </summary>
    ''' <param name="p_DJ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPdControlByDJ(ByVal p_DJ As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim DS As DataSet = New DataSet
            Try
                strSQL = String.Format("exec sp_121GetPdControlByDJ '{0}'", p_DJ)
                DS = da.ExecuteDataSet(strSQL, "T1")
            Catch ex As Exception
                ErrorLogging("121-GetPdControlByDJ", "MIC", ex.ToString)
            End Try
            Return DS
        End Using
    End Function

    ''' <summary>
    ''' Update POQty
    ''' add by michael
    ''' </summary>
    ''' <param name="PO"></param>
    ''' <param name="OrgCode"></param>
    ''' <param name="Model"></param>
    ''' <param name="ModelRev"></param>
    ''' <param name="POQty"></param>
    ''' <param name="TVA"></param>
    ''' <param name="AllowMatching"></param>
    ''' <param name="AllowPacking"></param>
    ''' <param name="ChangedBy"></param>
    ''' <param name="Remarks"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdatePOQty(ByVal PO As String, ByVal OrgCode As String, ByVal Model As String, ByVal ModelRev As String, ByVal POQty As Integer, ByVal TVA As String, ByVal AllowMatching As Boolean, ByVal AllowPacking As Boolean, ByVal ChangedBy As String, ByVal Remarks As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim strResult As String = ""
            Try
                strSQL = String.Format("exec sp_121UpdatePOQty '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}',N'{8}',N'{9}'", PO, OrgCode, Model, ModelRev, POQty, TVA, AllowMatching, AllowPacking, ChangedBy, Remarks)
                strResult = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                ErrorLogging("121-UpdatePOQty", "MIC", ex.ToString)
            End Try
            Return strResult
        End Using
    End Function

    Public Function ATELockingRDCWIPIN(ByVal IntSN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_ATELocking_RDCWIPIN '{0}'", IntSN)
                da.ExecuteNonQuery(strSQL)
                Return "1"
            Catch ex As Exception
                ErrorLogging("RDC-ATELockingRDCWIPIN", "", ex.Message & ex.Source, "E")
                Return "0"
            End Try
        End Using
    End Function
End Class
