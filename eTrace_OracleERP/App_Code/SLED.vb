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


Public Structure SelCriteria
    Public SType As String
    Public Matno As String
    Public ExpDate As Date
    Public RevStatus As String
    Public CLID As String
End Structure

Public Structure PostSLEDResponse
    Public Mess As String
    Public CLIDs As DataSet
    Public message As String
End Structure

Public Class SLED
    Inherits PublicFunction

    Public Function GetMatList(ByVal Selection As SelCriteria) As DataSet
        GetMatList = New DataSet
        Dim Materials As DataTable
        Dim myDataColumn As DataColumn
        Materials = New Data.DataTable("Materials")

        myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StorageBin", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Qty", System.Type.GetType("System.Double"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.DateTime"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("NextReviewDate", System.Type.GetType("System.DateTime"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SampleSize", System.Type.GetType("System.Int32"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReviewStatus", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReviewedOn", System.Type.GetType("System.DateTime"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReviewedBy", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OriNextReviewDate", System.Type.GetType("System.DateTime"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OriSampleSize", System.Type.GetType("System.Int32"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OriReviewStatus", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BaseUOM", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
        Materials.Columns.Add(myDataColumn)
        GetMatList.Tables.Add(Materials)

        Dim CLIDs As DataTable
        CLIDs = New Data.DataTable("CLIDs")

        myDataColumn = New Data.DataColumn("SLOC", System.Type.GetType("System.String"))
        CLIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StorageBin", System.Type.GetType("System.String"))
        CLIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        CLIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CLIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Qty", System.Type.GetType("System.Double"))
        CLIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SampleSize", System.Type.GetType("System.Int32"))
        CLIDs.Columns.Add(myDataColumn)

        GetMatList.Tables.Add(CLIDs)
        Dim myDataRow As Data.DataRow

        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim currMaterial As String = "", lastMaterial As String = ""
        Dim MatQty As Decimal = 0, MatSampleSize As Integer = 0
        Dim currItem As Integer

        If Selection.CLID Is Nothing Then
            If Selection.RevStatus = "" Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("SELECT CLID, MaterialNo, QtyBaseUOM, ExpDate, SLOC, StorageBin, NextReviewDate, ReviewStatus, ReviewedOn, ReviewedBy, SampleSize, RTLot,BaseUOM,MaterialRevision FROM T_CLMaster WHERE (StatusCode = 1 and SLOC Like @SLOC and MaterialNo Like @MaterialNo and ExpDate < @ExpDate and (NextReviewDate < @ExpDate or NextReviewDate is NULL) and ReviewStatus IS NULL) ORDER BY SLOC, StorageBin, MaterialNo, CLID", myConn)
            Else
                CLMasterSQLCommand = New SqlClient.SqlCommand("SELECT CLID, MaterialNo, QtyBaseUOM, ExpDate, SLOC, StorageBin, NextReviewDate, ReviewStatus, ReviewedOn, ReviewedBy, SampleSize, RTLot,BaseUOM,MaterialRevision FROM T_CLMaster WHERE (StatusCode = 1 and SLOC Like @SLOC and MaterialNo Like @MaterialNo and ExpDate < @ExpDate and (NextReviewDate < @ExpDate or NextReviewDate is NULL) and ReviewStatus like @ReviewStatus) ORDER BY SLOC, StorageBin, MaterialNo, CLID", myConn)
                CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
                CLMasterSQLCommand.Parameters("@ReviewStatus").Value = Selection.RevStatus
            End If
            CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 20, "SLOC")
            CLMasterSQLCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 50, "MaterialNo")
            CLMasterSQLCommand.Parameters.Add("@ExpDate", SqlDbType.DateTime, 4, "ExpDate")
            CLMasterSQLCommand.Parameters("@SLOC").Value = IIf(Selection.SType = "", "%", Selection.SType)
            CLMasterSQLCommand.Parameters("@MaterialNo").Value = IIf(Selection.Matno = "", "%", Selection.Matno)
            CLMasterSQLCommand.Parameters("@ExpDate").Value = Selection.ExpDate
        Else
            CLMasterSQLCommand = New SqlClient.SqlCommand("SELECT CLID, MaterialNo, QtyBaseUOM, ExpDate, SLOC, StorageBin, NextReviewDate, ReviewStatus, ReviewedOn, ReviewedBy, SampleSize, RTLot,BaseUOM,MaterialRevision FROM T_CLMaster WHERE (StatusCode = 1 and StorageBin IN (SELECT StorageBin FROM T_CLMaster WHERE CLID = @CLID) And MaterialNo IN (SELECT MaterialNo FROM T_CLMaster WHERE CLID = @CLID)) ORDER BY SLOC, StorageBin, MaterialNo, CLID", myConn)
            CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 25, "CLID")
            CLMasterSQLCommand.Parameters("@CLID").Value = Selection.CLID
        End If

        Try
            myConn.Open()
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                currMaterial = objReader.GetValue(4) & objReader.GetValue(5) & objReader.GetValue(1)
                If currMaterial <> lastMaterial Then
                    'Fill up the Qty and Sample Size
                    currItem = Materials.Rows.Count - 1
                    If currItem >= 0 Then
                        Materials.Rows(currItem).Item("Qty") = MatQty
                        Materials.Rows(currItem).Item("SampleSize") = MatSampleSize
                        Materials.Rows(currItem).Item("OriSampleSize") = MatSampleSize
                        Materials.Rows(currItem).AcceptChanges()
                    End If
                    myDataRow = Materials.NewRow()
                    If Not objReader.GetValue(1) Is DBNull.Value Then myDataRow("MaterialNo") = objReader.GetValue(1)
                    If Not objReader.GetValue(3) Is DBNull.Value Then myDataRow("ExpDate") = objReader.GetValue(3)
                    If Not objReader.GetValue(4) Is DBNull.Value Then myDataRow("SLOC") = objReader.GetValue(4)
                    If Not objReader.GetValue(5) Is DBNull.Value Then myDataRow("StorageBin") = objReader.GetValue(5)
                    If Not objReader.GetValue(6) Is DBNull.Value Then myDataRow("NextReviewDate") = objReader.GetValue(6)
                    If Not objReader.GetValue(7) Is DBNull.Value Then myDataRow("ReviewStatus") = objReader.GetValue(7)
                    If Not objReader.GetValue(8) Is DBNull.Value Then myDataRow("ReviewedOn") = objReader.GetValue(8)
                    If Not objReader.GetValue(9) Is DBNull.Value Then myDataRow("ReviewedBy") = objReader.GetValue(9)
                    If Not objReader.GetValue(6) Is DBNull.Value Then myDataRow("OriNextReviewDate") = objReader.GetValue(6)
                    If Not objReader.GetValue(7) Is DBNull.Value Then myDataRow("OriReviewStatus") = objReader.GetValue(7)
                    If Not objReader.GetValue(11) Is DBNull.Value Then myDataRow("RTLot") = objReader.GetValue(11)
                    If Not objReader.GetValue(12) Is DBNull.Value Then myDataRow("BaseUOM") = objReader.GetValue(12)
                    If Not objReader.GetValue(13) Is DBNull.Value Then myDataRow("Revision") = objReader.GetValue(13)
                    Materials.Rows.Add(myDataRow)

                    lastMaterial = currMaterial
                    MatQty = 0
                    MatSampleSize = 0
                End If
                myDataRow = CLIDs.NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then myDataRow("CLID") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then myDataRow("MaterialNo") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then myDataRow("Qty") = objReader.GetValue(2)
                If Not objReader.GetValue(4) Is DBNull.Value Then myDataRow("SLOC") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then myDataRow("StorageBin") = objReader.GetValue(5)
                If Not objReader.GetValue(10) Is DBNull.Value Then myDataRow("SampleSize") = objReader.GetValue(10)
                CLIDs.Rows.Add(myDataRow)
                MatQty = MatQty + objReader.GetValue(2)
                MatSampleSize = MatSampleSize + IIf(objReader.GetValue(10) Is DBNull.Value, 0, objReader.GetValue(10))
            End While
            'Fill up the Qty and Sample Size for last item
            currItem = Materials.Rows.Count - 1
            If currItem >= 0 Then
                Materials.Rows(currItem).Item("Qty") = MatQty
                Materials.Rows(currItem).Item("SampleSize") = MatSampleSize
                Materials.Rows(currItem).Item("OriSampleSize") = MatSampleSize
                Materials.Rows(currItem).AcceptChanges()
            End If
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("SLED_GetMatList", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function


    Public Function SLEDSaveChanges(ByVal OracleLoginData As ERPLogin, ByVal CLIDs As DataSet) As PostSLEDResponse

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpdateSAP As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLoc=@SLoc,StorageType=@StorageType,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLoc", SqlDbType.VarChar, 50, "SLoc")
        CLMasterSQLCommand.Parameters.Add("@StorageType", SqlDbType.VarChar, 20, "StorageType")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpdateSAP = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))
                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then
                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        If QtyDiff > 0 Then
                            'Call SAP RFC to update SAP inventory. Input: MatNo, Stype, Sbin, QtyDiff
                            'GMHeader.Doc_Date = DateTime.Now.ToString("yyyyMMdd")
                            'GMHeader.Pstng_Date = DateTime.Now.ToString("yyyyMMdd")
                            ''GMHeader.Header_Txt = 
                            'GMCode.Gm_Code = "03"       'For Scrap Material or Consumption Material

                            'GMItem = New BAPI2017_GM_ITEM_CREATE
                            'GMItem.Material = CLIDs.Tables(0).Rows(i)("MaterialNo")
                            'GMItem.Plant = "3110"
                            'GMItem.Stge_Loc = "0010"
                            'GMItem.Move_Type = "551"
                            'GMItem.Entry_Qnt = QtyDiff
                            'GMItem.Gr_Rcpt = ""
                            'GMItem.Costcenter = "0000031640"
                            'GMItem.No_Transfer_Req = "X"
                            'GMItem.Stge_Type = CLIDs.Tables(0).Rows(i)("StorageType")
                            'GMItem.Stge_Bin = CLIDs.Tables(0).Rows(i)("StorageBin")
                            'GMItems.Add(GMItem)

                            UpdateSAP = True

                        ElseIf QtyDiff < 0 Then
                            UpdateSAP = False
                        End If
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()
                    ElseIf CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then

                        'GMHeader.Doc_Date = DateTime.Now.ToString("yyyyMMdd")
                        'GMHeader.Pstng_Date = DateTime.Now.ToString("yyyyMMdd")
                        'GMCode.Gm_Code = "04"       'For Transfer Posting

                        'GMItem = New BAPI2017_GM_ITEM_CREATE
                        'GMItem.Material = CLIDs.Tables(0).Rows(i)("MaterialNo")
                        'GMItem.Plant = "3110"
                        'GMItem.Stge_Loc = "0010"
                        'GMItem.Move_Type = "311"
                        'GMItem.Entry_Qnt = CLIDs.Tables(0).Rows(i)("Qty")
                        'GMItem.Gr_Rcpt = ""
                        'GMItem.Move_Stloc = "9995"
                        'GMItem.No_Transfer_Req = "X"
                        'GMItem.Stge_Type = CLIDs.Tables(0).Rows(i)("StorageType")
                        'GMItem.Stge_Bin = CLIDs.Tables(0).Rows(i)("StorageBin")
                        'GMItems.Add(GMItem)

                        UpdateSAP = True

                    ElseIf CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Passed" Then
                        UpdateSAP = False
                    End If

                    If UpdateSAP = True Then
                        'Dim mvtHeaderRet As BAPI2017_GM_HEAD_RET = New BAPI2017_GM_HEAD_RET
                        '   Dim MatDocNo As String
                        '  Dim MatDocYear As String
                        'Dim GMSerialNo As BAPI2017_GM_SERIALNUMBERTable = New BAPI2017_GM_SERIALNUMBERTable
                        'Dim GMBAPIRet As BAPIRET2Table = New BAPIRET2Table
                        'Dim BAPIRet As BAPIRET2 = New BAPIRET2

                        'IQCCreateProxy = New SAPIQCProxy(SAPConnectionStr)
                        'IQCCreateProxy.Connection.Open()
                        'IQCCreateProxy.Z_Etrace_Goodsmvt_Create(GMCode, GMHeader, GMTestRun, mvtHeaderRet, MatDocYear, MatDocNo, GMItems, GMSerialNo, GMBAPIRet)

                        'IQCCreateProxy.Connection.Close()

                        'For j = 0 To GMBAPIRet.Count - 1
                        'SLEDSaveChanges.Message = SLEDSaveChanges.Message & GMBAPIRet.Item(j).Message.ToString()
                        'Next

                        ' SLEDSaveChanges.DocNo = MatDocNo
                        ' SLEDSaveChanges.DocYear = MatDocYear

                        'If MatDocNo = "" Then
                        '    SLEDSaveChanges.Message = "^IQC-26@" & SLEDSaveChanges.Message
                        '    Exit Function
                        'End If
                    End If

                    DR = CLIDs.Tables(1).Select("StorageType = '" & CLIDs.Tables(0).Rows(i)("StorageType") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")
                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))
                        CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")
                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        ' CLMasterSQLCommand.Parameters("@ReviewedBy").Value = SAPLoginData.User.ToUpper
                        'Change SLOC to 9995 for failed CLIDs
                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then
                            CLMasterSQLCommand.Parameters("@SLoc").Value = "9995"
                            CLMasterSQLCommand.Parameters("@StorageType").Value = ""
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = ""
                        Else
                            CLMasterSQLCommand.Parameters("@SLoc").Value = "0010"
                            CLMasterSQLCommand.Parameters("@StorageType").Value = CLIDs.Tables(0).Rows(i)("StorageType")
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")
                        End If

                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()
            SLEDSaveChanges.CLIDs = CLIDs
        Catch ex As Exception
            SLEDSaveChanges.message = "^IQC-26@ " & ex.Message
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    ' ''Public Function SLED_Issue(ByVal p_ds As DataSet, ByVal CLIDs As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet

    ' ''    Dim i, j As Integer
    ' ''    Dim CLMasterSQLCommand As SqlClient.SqlCommand
    ' ''    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
    ' ''    Dim ra As Integer
    ' ''    Dim QtyDiff As Decimal
    ' ''    Dim DR() As DataRow
    ' ''    Dim NRDate, OriNRDate, nullDate As DateTime
    ' ''    Dim UpOracle As Boolean

    ' ''    CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLoc=@SLoc,StorageBin=@StorageBin where CLID=@CLID ", myConn)
    ' ''    CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
    ' ''    CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
    ' ''    CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
    ' ''    CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
    ' ''    CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
    ' ''    CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
    ' ''    CLMasterSQLCommand.Parameters.Add("@SLoc", SqlDbType.VarChar, 50, "SLoc")
    ' ''    CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")


    ' ''    Try
    ' ''        myConn.Open()
    ' ''        For i = 0 To CLIDs.Tables(0).Rows.Count - 1
    ' ''            UpOracle = False
    ' ''            If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
    ' ''            If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
    ' ''            NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
    ' ''            OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

    ' ''            If NRDate <> OriNRDate _
    ' ''                Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
    ' ''                Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

    ' ''                If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then
    ' ''                    QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
    ' ''                    If QtyDiff > 0 Then
    ' ''                        UpOracle = True
    ' ''                    ElseIf QtyDiff < 0 Then
    ' ''                        UpOracle = False
    ' ''                    End If

    ' ''                    CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
    ' ''                    CLIDs.Tables(0).Rows(i).AcceptChanges()
    ' ''                ElseIf CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then
    ' ''                    UpOracle = True
    ' ''                ElseIf CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Passed" Then
    ' ''                    UpOracle = False
    ' ''                End If

    ' ''                If UpOracle = True Then
    ' ''                    Dim oda_h As New OracleDataAdapter()
    ' ''                    Dim qcomm As New OracleCommand()
    ' ''                    Dim oc As New OracleConnection("Data Source=CAROLD3;user id=APPS;password=apps")

    ' ''                    qcomm.Connection = oc
    ' ''                    qcomm.CommandType = CommandType.StoredProcedure
    ' ''                    qcomm.CommandText = "apps.cux_inv_mtl_tran_pkg.account_alias_issue"
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_organization_code", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_transaction_source_name", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_transaction_uom", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_source_line_id", OracleType.Int32))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_source_header_id", OracleType.Int32))

    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))

    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_item_segment1", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_item_revision", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_subinventory_code", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_loc_segment1", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_loc_segment2", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_loc_segment3", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_loc_segment4", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_loc_segment5", OracleType.VarChar, 240))

    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_lot_expiration_date", OracleType.DateTime))

    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_transaction_quantity", OracleType.Int32))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_primary_quantity", OracleType.Int32))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_reason_code", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("p_transaction_reference", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("o_return_status", OracleType.VarChar, 240))
    ' ''                    qcomm.Parameters.Add(New OracleParameter("o_return_message", OracleType.VarChar, 240))

    ' ''                    qcomm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
    ' ''                    qcomm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

    ' ''                    qcomm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
    ' ''                    qcomm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
    ' ''                    qcomm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
    ' ''                    qcomm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
    ' ''                    qcomm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
    ' ''                    qcomm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
    ' ''                    qcomm.Parameters("p_user_id").SourceColumn = "p_user_id"
    ' ''                    qcomm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
    ' ''                    qcomm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
    ' ''                    qcomm.Parameters("p_subinventory_code").SourceColumn = "p_subinventory_code"
    ' ''                    qcomm.Parameters("p_loc_segment1").SourceColumn = "p_loc_segment1"
    ' ''                    qcomm.Parameters("p_loc_segment2").SourceColumn = "p_loc_segment2"
    ' ''                    qcomm.Parameters("p_loc_segment3").SourceColumn = "p_loc_segment3"
    ' ''                    qcomm.Parameters("p_loc_segment4").SourceColumn = "p_loc_segment4"
    ' ''                    qcomm.Parameters("p_loc_segment5").SourceColumn = "p_loc_segment5"

    ' ''                    qcomm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
    ' ''                    qcomm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"

    ' ''                    qcomm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
    ' ''                    qcomm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
    ' ''                    qcomm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
    ' ''                    qcomm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
    ' ''                    qcomm.Parameters("o_return_status").SourceColumn = "o_return_status"
    ' ''                    qcomm.Parameters("o_return_message").SourceColumn = "o_return_message"

    ' ''                    qcomm.Connection.Open()
    ' ''                    oda_h.InsertCommand = qcomm
    ' ''                    oda_h.Update(p_ds.Tables("SLEDissue_table"))

    ' ''                    qcomm.Connection.Close()


    ' ''                    ErrorLogging("SLED", "Yong2", p_ds.Tables("SLEDissue_table").Rows(0)("o_return_status"))
    ' ''                    ErrorLogging("SLED", "Yong2", p_ds.Tables("SLEDissue_table").Rows(0)("o_return_message"))

    ' ''                    If p_ds.Tables("SLEDissue_table").Rows(0)("o_return_status") = "Y" Then
    ' ''                        Return transfer_sled_submit("SLED", OracleLoginData, TransactionID, 1800000)
    ' ''                    End If
    ' ''                End If


    ' ''                ''''''''''''''''''''''''''
    ' ''                If p_ds.Tables("SLEDissue_table").Rows(0)("o_return_status") = "Y" Then

    ' ''                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
    ' ''                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
    ' ''                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

    ' ''                    For j = 0 To DR.Length - 1
    ' ''                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
    ' ''                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))
    ' ''                        CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
    ' ''                        CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")
    ' ''                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
    ' ''                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

    ' ''                        'Change SLOC to 9995 for failed CLIDs
    ' ''                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then
    ' ''                            CLMasterSQLCommand.Parameters("@SLoc").Value = "9995"
    ' ''                            ' CLMasterSQLCommand.Parameters("@StorageType").Value = ""
    ' ''                            CLMasterSQLCommand.Parameters("@StorageBin").Value = ""
    ' ''                        Else
    ' ''                            CLMasterSQLCommand.Parameters("@SLoc").Value = "0010"
    ' ''                            ' CLMasterSQLCommand.Parameters("@StorageType").Value = CLIDs.Tables(0).Rows(i)("StorageType")
    ' ''                            CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")
    ' ''                        End If

    ' ''                        CLMasterSQLCommand.CommandType = CommandType.Text
    ' ''                        ra = CLMasterSQLCommand.ExecuteNonQuery()

    ' ''                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
    ' ''                        DR(j).AcceptChanges()
    ' ''                    Next
    ' ''                End If
    ' ''            End If
    ' ''        Next
    ' ''        myConn.Close()
    ' ''        SLED_Issue = CLIDs

    ' ''    Catch ex As Exception
    ' ''        ErrorLogging("SLED-Post", "", ex.Message & ex.Source, "E")
    ' ''    End Try

    ' ''End Function

    ' ''Public Function SLED_ReviewUpd(ByVal p_ds As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet

    ' ''    Try

    ' ''        Dim oda_h As New OracleDataAdapter()
    ' ''        Dim qcomm As New OracleCommand()
    ' ''        Dim oc As New OracleConnection("Data Source=CAROLD3;user id=APPS;password=apps")

    ' ''        qcomm.Connection = oc
    ' ''        qcomm.CommandType = CommandType.StoredProcedure
    ' ''        qcomm.CommandText = "apps.cux_inv_mtl_tran_pkg.account_alias_issue"
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_organization_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_source_name", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_uom", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_source_line_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_source_header_id", OracleType.Int32))

    ' ''        qcomm.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))

    ' ''        qcomm.Parameters.Add(New OracleParameter("p_item_segment1", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_item_revision", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_subinventory_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment1", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment2", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment3", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment4", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment5", OracleType.VarChar, 240))

    ' ''        qcomm.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_lot_expiration_date", OracleType.DateTime))

    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_quantity", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_primary_quantity", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_reason_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_reference", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("o_return_status", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("o_return_message", OracleType.VarChar, 240))

    ' ''        qcomm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
    ' ''        qcomm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

    ' ''        qcomm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
    ' ''        qcomm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
    ' ''        qcomm.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
    ' ''        qcomm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
    ' ''        qcomm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
    ' ''        qcomm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
    ' ''        qcomm.Parameters("p_user_id").SourceColumn = "p_user_id"
    ' ''        qcomm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
    ' ''        qcomm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
    ' ''        qcomm.Parameters("p_subinventory_code").SourceColumn = "p_subinventory_code"
    ' ''        qcomm.Parameters("p_loc_segment1").SourceColumn = "p_loc_segment1"
    ' ''        qcomm.Parameters("p_loc_segment2").SourceColumn = "p_loc_segment2"
    ' ''        qcomm.Parameters("p_loc_segment3").SourceColumn = "p_loc_segment3"
    ' ''        qcomm.Parameters("p_loc_segment4").SourceColumn = "p_loc_segment4"
    ' ''        qcomm.Parameters("p_loc_segment5").SourceColumn = "p_loc_segment5"

    ' ''        qcomm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
    ' ''        qcomm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"

    ' ''        qcomm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
    ' ''        qcomm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
    ' ''        qcomm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
    ' ''        qcomm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
    ' ''        qcomm.Parameters("o_return_status").SourceColumn = "o_return_status"
    ' ''        qcomm.Parameters("o_return_message").SourceColumn = "o_return_message"

    ' ''        qcomm.Connection.Open()
    ' ''        oda_h.InsertCommand = qcomm
    ' ''        oda_h.Update(p_ds.Tables("SLEDissue_table"))

    ' ''        qcomm.Connection.Close()

    ' ''        'ErrorLogging("SLED_review", "Yong2", p_ds.Tables("SLEDissue_table").Rows(0)("o_return_status"))
    ' ''        'ErrorLogging("SLED_review", "Yong2", p_ds.Tables("SLEDissue_table").Rows(0)("o_return_message"))


    ' ''        Return transfer_sled_submit("SLED", OracleLoginData, TransactionID, 1800000)

    ' ''    Catch oe As Exception
    ' ''        ErrorLogging("SLED", OracleLoginData, oe.Message & oe.Source, "E")
    ' ''        Throw oe
    ' ''    End Try

    ' ''End Function



    ' ''Public Function SLED_FailUpd(ByVal p_ds_fail As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet
    ' ''    Dim i As Integer
    ' ''    Try

    ' ''        Dim oda_h As New OracleDataAdapter()
    ' ''        Dim qcomm As New OracleCommand()
    ' ''        Dim oc As New OracleConnection("Data Source=CAROLD3;user id=APPS;password=apps")

    ' ''        qcomm.Connection = oc
    ' ''        qcomm.CommandType = CommandType.StoredProcedure
    ' ''        qcomm.CommandText = "apps.cux_inv_mtl_tran_pkg.subinventory_transfer"
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_organization_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_uom", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_source_line_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_source_header_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_item_segment1", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_item_revision", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transfer_subinventory", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_xfer_loc_segment1", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_xfer_loc_segment2", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_xfer_loc_segment3", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_xfer_loc_segment4", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_xfer_loc_segment5", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_subinventory_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment1", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment2", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment3", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment4", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_loc_segment5", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_lot_expiration_date", OracleType.DateTime))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_quantity", OracleType.Double))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_primary_quantity", OracleType.Double))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_reason_code", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("p_transaction_reference", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("o_return_status", OracleType.VarChar, 240))
    ' ''        qcomm.Parameters.Add(New OracleParameter("o_return_message", OracleType.VarChar, 240))

    ' ''        qcomm.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
    ' ''        qcomm.Parameters("o_return_message").Direction = ParameterDirection.InputOutput
    ' ''        qcomm.Parameters("p_organization_code").SourceColumn = "p_organization_code"
    ' ''        qcomm.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
    ' ''        qcomm.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
    ' ''        qcomm.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
    ' ''        qcomm.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
    ' ''        qcomm.Parameters("p_user_id").SourceColumn = "p_user_id"
    ' ''        qcomm.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
    ' ''        qcomm.Parameters("p_item_revision").SourceColumn = "p_item_revision"
    ' ''        qcomm.Parameters("p_transfer_subinventory").SourceColumn = "p_transfer_subinventory"
    ' ''        qcomm.Parameters("p_xfer_loc_segment1").SourceColumn = "p_xfer_loc_segment1"
    ' ''        qcomm.Parameters("p_xfer_loc_segment2").SourceColumn = "p_xfer_loc_segment2"
    ' ''        qcomm.Parameters("p_xfer_loc_segment3").SourceColumn = "p_xfer_loc_segment3"
    ' ''        qcomm.Parameters("p_xfer_loc_segment4").SourceColumn = "p_xfer_loc_segment4"
    ' ''        qcomm.Parameters("p_xfer_loc_segment5").SourceColumn = "p_xfer_loc_segment5"
    ' ''        qcomm.Parameters("p_subinventory_code").SourceColumn = "p_subinventory_code"
    ' ''        qcomm.Parameters("p_loc_segment1").SourceColumn = "p_loc_segment1"
    ' ''        qcomm.Parameters("p_loc_segment2").SourceColumn = "p_loc_segment2"
    ' ''        qcomm.Parameters("p_loc_segment3").SourceColumn = "p_loc_segment3"
    ' ''        qcomm.Parameters("p_loc_segment4").SourceColumn = "p_loc_segment4"
    ' ''        qcomm.Parameters("p_loc_segment5").SourceColumn = "p_loc_segment5"
    ' ''        qcomm.Parameters("p_lot_number").SourceColumn = "p_lot_number"
    ' ''        qcomm.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
    ' ''        qcomm.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
    ' ''        qcomm.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
    ' ''        qcomm.Parameters("p_reason_code").SourceColumn = "p_reason_code"
    ' ''        qcomm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
    ' ''        qcomm.Parameters("o_return_status").SourceColumn = "o_return_status"
    ' ''        qcomm.Parameters("o_return_message").SourceColumn = "o_return_message"

    ' ''        qcomm.Connection.Open()
    ' ''        oda_h.InsertCommand = qcomm
    ' ''        oda_h.Update(p_ds_fail.Tables("fail_table"))
    ' ''        qcomm.Connection.Close()

    ' ''        ErrorLogging("SLED_Fail", "Yong2", p_ds_fail.Tables("fail_table").Rows(0)("o_return_status"))
    ' ''        ErrorLogging("SLED_Fail", "Yong2", p_ds_fail.Tables("fail_table").Rows(0)("o_return_message"))

    ' ''        For i = 0 To p_ds_fail.Tables("fail_table").Rows.Count - 1
    ' ''            If p_ds_fail.Tables("fail_table").Rows(i)("o_return_status") = "N" Then
    ' ''                Return p_ds_fail
    ' ''                Exit Function
    ' ''            End If
    ' ''        Next

    ' ''        Return transfer_sled_submit("SLED_FailUpd", OracleLoginData, TransactionID, 1800000)

    ' ''    Catch oe As Exception
    ' ''        ErrorLogging("SLED_FailUpd", OracleLoginData, oe.Message & oe.Source, "E")
    ' ''        Throw oe
    ' ''    End Try

    ' ''End Function


    Public Function SLED_ReviewUpd(ByVal p_ds As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim oda_h As OracleDataAdapter = da.Oda_Insert_Tran

            Try
                oda_h.InsertCommand.CommandType = CommandType.StoredProcedure
                oda_h.InsertCommand.CommandText = "apps.cux_inv_mtl_tran_pkg.account_alias_issue"
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_timeout", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_organization_code", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_source_name", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_uom", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_source_line_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_source_header_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_item_segment1", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_item_revision", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_subinventory_source", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_locator_source", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_lot_expiration_date", OracleType.DateTime))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_quantity", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_primary_quantity", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_reason_code", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_reference", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("o_return_status", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("o_return_message", OracleType.VarChar, 240))

                oda_h.InsertCommand.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
                oda_h.InsertCommand.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

                oda_h.InsertCommand.Parameters("p_timeout").SourceColumn = "p_timeout"
                oda_h.InsertCommand.Parameters("p_organization_code").SourceColumn = "p_organization_code"
                oda_h.InsertCommand.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
                oda_h.InsertCommand.Parameters("p_transaction_source_name").SourceColumn = "p_transaction_source_name"
                oda_h.InsertCommand.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
                oda_h.InsertCommand.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
                oda_h.InsertCommand.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
                oda_h.InsertCommand.Parameters("p_user_id").SourceColumn = "p_user_id"
                oda_h.InsertCommand.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
                oda_h.InsertCommand.Parameters("p_item_revision").SourceColumn = "p_item_revision"
                oda_h.InsertCommand.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
                oda_h.InsertCommand.Parameters("p_locator_source").SourceColumn = "p_locator_source"
                oda_h.InsertCommand.Parameters("p_lot_number").SourceColumn = "p_lot_number"
                oda_h.InsertCommand.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
                oda_h.InsertCommand.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
                oda_h.InsertCommand.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
                oda_h.InsertCommand.Parameters("p_reason_code").SourceColumn = "p_reason_code"
                oda_h.InsertCommand.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
                oda_h.InsertCommand.Parameters("o_return_status").SourceColumn = "o_return_status"
                oda_h.InsertCommand.Parameters("o_return_message").SourceColumn = "o_return_message"

                oda_h.Update(p_ds.Tables("SLEDissue_table"))
                Dim DR() As DataRow = Nothing

                Dim i As Integer
                DR = p_ds.Tables("SLEDissue_table").Select("o_return_status = 'N'")

                If DR.Length = 0 Then
                    oda_h.InsertCommand.Transaction.Commit()
                    oda_h.InsertCommand.Connection.Close()
                    Return p_ds
                Else
                    For i = 0 To (DR.Length - 1)
                        ErrorLogging("SLED-Issue", OracleLoginData, DR(i)("o_return_message").ToString)
                    Next
                    oda_h.InsertCommand.Transaction.Rollback()
                    oda_h.InsertCommand.Connection.Close()
                    Return p_ds
                End If

            Catch ex As Exception
                oda_h.InsertCommand.Transaction.Rollback()
                Return p_ds
            Finally
                If oda_h.InsertCommand.Connection.State <> ConnectionState.Closed Then oda_h.InsertCommand.Connection.Close()
            End Try
        End Using

    End Function


    Public Function SLED_FailUpd(ByVal p_ds_fail As DataSet, ByVal OracleLoginData As String, ByVal TransactionID As Integer) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim oda_h As OracleDataAdapter = da.Oda_Insert_Tran

            Try
                oda_h.InsertCommand.CommandType = CommandType.StoredProcedure
                oda_h.InsertCommand.CommandText = "apps.cux_inv_mtl_tran_pkg.subinventory_transfer"
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_timeout", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_organization_code", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_header_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_uom", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_source_line_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_source_header_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_user_id", OracleType.Int32))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_item_segment1", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_item_revision", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_subinventory_source", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_locator_source", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_subinventory_destination", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_locator_destination", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_lot_number", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_lot_expiration_date", OracleType.DateTime))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_quantity", OracleType.Double))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_primary_quantity", OracleType.Double))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_reason_code", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("p_transaction_reference", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("o_return_status", OracleType.VarChar, 240))
                oda_h.InsertCommand.Parameters.Add(New OracleParameter("o_return_message", OracleType.VarChar, 240))

                oda_h.InsertCommand.Parameters("p_subinventory_destination").Direction = ParameterDirection.InputOutput
                oda_h.InsertCommand.Parameters("o_return_status").Direction = ParameterDirection.InputOutput
                oda_h.InsertCommand.Parameters("o_return_message").Direction = ParameterDirection.InputOutput

                oda_h.InsertCommand.Parameters("p_timeout").SourceColumn = "p_timeout"
                oda_h.InsertCommand.Parameters("p_organization_code").SourceColumn = "p_organization_code"
                oda_h.InsertCommand.Parameters("p_transaction_header_id").SourceColumn = "p_transaction_header_id"
                oda_h.InsertCommand.Parameters("p_transaction_uom").SourceColumn = "p_transaction_uom"
                oda_h.InsertCommand.Parameters("p_source_line_id").SourceColumn = "p_source_line_id"
                oda_h.InsertCommand.Parameters("p_source_header_id").SourceColumn = "p_source_header_id"
                oda_h.InsertCommand.Parameters("p_user_id").SourceColumn = "p_user_id"
                oda_h.InsertCommand.Parameters("p_item_segment1").SourceColumn = "p_item_segment1"
                oda_h.InsertCommand.Parameters("p_item_revision").SourceColumn = "p_item_revision"
                oda_h.InsertCommand.Parameters("p_subinventory_source").SourceColumn = "p_subinventory_source"
                oda_h.InsertCommand.Parameters("p_locator_source").SourceColumn = "p_locator_source"
                oda_h.InsertCommand.Parameters("p_subinventory_destination").SourceColumn = "p_subinventory_destination"
                oda_h.InsertCommand.Parameters("p_locator_destination").SourceColumn = "p_locator_destination"
                oda_h.InsertCommand.Parameters("p_lot_number").SourceColumn = "p_lot_number"
                oda_h.InsertCommand.Parameters("p_lot_expiration_date").SourceColumn = "p_lot_expiration_date"
                oda_h.InsertCommand.Parameters("p_transaction_quantity").SourceColumn = "p_transaction_quantity"
                oda_h.InsertCommand.Parameters("p_primary_quantity").SourceColumn = "p_primary_quantity"
                oda_h.InsertCommand.Parameters("p_reason_code").SourceColumn = "p_reason_code"
                oda_h.InsertCommand.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
                oda_h.InsertCommand.Parameters("o_return_status").SourceColumn = "o_return_status"
                oda_h.InsertCommand.Parameters("o_return_message").SourceColumn = "o_return_message"

                oda_h.Update(p_ds_fail.Tables("fail_table"))
                Dim DR() As DataRow = Nothing

                Dim i As Integer
                DR = p_ds_fail.Tables("fail_table").Select("o_return_status = 'N'")

                If DR.Length = 0 Then
                    oda_h.InsertCommand.Transaction.Commit()
                    oda_h.InsertCommand.Connection.Close()
                    Return p_ds_fail
                Else
                    For i = 0 To (DR.Length - 1)
                        ErrorLogging("SLED", OracleLoginData, DR(i)("o_return_message").ToString)
                    Next
                    oda_h.InsertCommand.Transaction.Rollback()
                    oda_h.InsertCommand.Connection.Close()
                    Return p_ds_fail
                End If

            Catch ex As Exception
                oda_h.InsertCommand.Transaction.Rollback()
                Return p_ds_fail
            Finally
                If oda_h.InsertCommand.Connection.State <> ConnectionState.Closed Then oda_h.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function


    Public Function transfer_sled_submit(ByVal p_type As String, ByVal p_OracleLoginData As String, ByVal p_transaction_header_id As Integer, ByVal p_timeout As Integer) As DataSet
        Dim p_ds As New DataSet()
        p_ds.Tables.Add("MSG")

        Dim oda As New OracleDataAdapter()
        Dim comm_submit As New OracleCommand()
        Dim oc As New OracleConnection("Data Source=CAROLD3;user id=APPS;password=apps")

        Try

            comm_submit.Connection = oc
            oda.SelectCommand = comm_submit
            oda.SelectCommand.CommandType = CommandType.StoredProcedure

            oda.SelectCommand.CommandText = "apps.cux_inv_mtl_tran_pkg.Call_Process_OnLine"

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

            'ErrorLogging("p_tran_result", "Yong2", oda.SelectCommand.Parameters("p_tran_result").Value.ToString.ToUpper)
            'ErrorLogging("p_error_code", "Yong2", oda.SelectCommand.Parameters("p_error_code").Value.ToString.ToUpper)
            ErrorLogging("p_error_explanation", "Yong2", oda.SelectCommand.Parameters("p_error_explanation").Value.ToString.ToUpper)
            ErrorLogging("p_process_online_yn", "Yong2", oda.SelectCommand.Parameters("p_process_online_yn").Value.ToString.ToUpper)
            'ErrorLogging("x_return_status", "Yong", oda.SelectCommand.Parameters("x_return_status").Value.ToString.ToUpper)
            'ErrorLogging("x_msg_count", "Yong", oda.SelectCommand.Parameters("x_msg_count").Value.ToString.ToUpper)
            'ErrorLogging("x_msg_data", "Yong", oda.SelectCommand.Parameters("x_msg_data").Value.ToString.ToUpper)
            'ErrorLogging("x_trans_count", "Yong", oda.SelectCommand.Parameters("x_trans_count").Value.ToString.ToUpper)
            'ErrorLogging("p_error_message", "Yong", oda.SelectCommand.Parameters("p_error_message").Value.ToString.ToUpper)

        Catch oe As Exception
            ErrorLogging(p_type, p_OracleLoginData, oe.Message & oe.Source, "E")
            Throw oe
        Finally
            If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
        End Try
        Return p_ds
    End Function


    Public Function eTrace_Upd(ByVal CLIDs As DataSet, ByVal OracleLoginData As String) As PostSLEDResponse

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpOracle As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLOC=@SLOC,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpOracle = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then

                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()

                    End If

                    ''''''''''''''''''''''''''
                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))
                        CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Passed" Then
                            CLMasterSQLCommand.Parameters("@ReviewStatus").Value = System.DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")
                        End If
                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then
                            CLMasterSQLCommand.Parameters("@SLOC").Value = "ZQ0 MRBR"
                            'CLMasterSQLCommand.Parameters("@StorageType").Value = ""
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = ""
                        Else
                            CLMasterSQLCommand.Parameters("@SLOC").Value = CLIDs.Tables(0).Rows(i)("SLOC")
                            'CLMasterSQLCommand.Parameters("@StorageType").Value = CLIDs.Tables(0).Rows(i)("StorageType")
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")
                        End If

                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()
            eTrace_Upd.CLIDs = CLIDs
        Catch ex As Exception
            eTrace_Upd.Mess = "^IQC-26@ " & ex.Message
            ErrorLogging("SLED-eTraceUpdt", OracleLoginData, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function


    Public Function SLEDSaveTrace(ByVal CLIDs As DataSet, ByVal OracleLoginData As String) As PostSLEDResponse

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpOracle As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLOC=@SLOC,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpOracle = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then

                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()

                    End If

                    ''''''''''''''''''''''''''
                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))

                        If CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value Then
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        End If

                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Passed" Then
                            CLMasterSQLCommand.Parameters("@ReviewStatus").Value = System.DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")
                        End If

                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

                        If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Failed" Then
                            CLMasterSQLCommand.Parameters("@SLOC").Value = "ZQ0 MRBR"
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = ""
                        Else
                            CLMasterSQLCommand.Parameters("@SLOC").Value = CLIDs.Tables(0).Rows(i)("SLOC")
                            CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")
                        End If

                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()
            SLEDSaveTrace.CLIDs = CLIDs
        Catch ex As Exception
            SLEDSaveTrace.Mess = ex.Message
            ErrorLogging("SLED-eTraceUpdt", OracleLoginData, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function


    Public Function eTraceUpd_Review(ByVal CLIDs As DataSet, ByVal OracleLoginData As String) As String

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpOracle As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLOC=@SLOC,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpOracle = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then

                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()

                    End If

                    ''''''''''''''''''''''''''
                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))

                        If CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value Then
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        End If

                        CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")

                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

                        CLMasterSQLCommand.Parameters("@SLOC").Value = CLIDs.Tables(0).Rows(i)("SLOC")
                        CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")

                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()

        Catch ex As Exception
            eTraceUpd_Review = ex.Message
            ErrorLogging("SLED-eTraceUpd_Review", OracleLoginData, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function eTraceUpd_Fail(ByVal CLIDs As DataSet, ByVal OracleLoginData As String) As String

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpOracle As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLOC=@SLOC,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpOracle = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then

                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()

                    End If

                    ''''''''''''''''''''''''''
                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))

                        If CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value Then
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        End If

                        CLMasterSQLCommand.Parameters("@ReviewStatus").Value = CLIDs.Tables(0).Rows(i)("ReviewStatus")

                        If DR(j)("SampleSize") Is DBNull.Value Then
                            DR(j)("SampleSize") = 0
                        End If

                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

                        CLMasterSQLCommand.Parameters("@SLOC").Value = "ZQ0 MRBR"
                        CLMasterSQLCommand.Parameters("@StorageBin").Value = ""


                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()
        Catch ex As Exception
            eTraceUpd_Fail = ex.Message
            ErrorLogging("SLED-eTraceUpd_Fail", OracleLoginData, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function eTraceUpd_Pass(ByVal CLIDs As DataSet, ByVal OracleLoginData As String) As String

        Dim i, j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim QtyDiff As Decimal
        Dim DR() As DataRow
        Dim NRDate, OriNRDate, nullDate As DateTime
        Dim UpOracle As Boolean

        CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set QtyBaseUOM=@Qty,NextReviewDate=@NextReviewDate,ReviewStatus=@ReviewStatus,SampleSize=@SampleSize,ReviewedOn=getdate(),ReviewedBy=@ReviewedBy,SLOC=@SLOC,StorageBin=@StorageBin where CLID=@CLID ", myConn)
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
        CLMasterSQLCommand.Parameters.Add("@NextReviewDate", SqlDbType.SmallDateTime, 4, "NextReviewDate")
        CLMasterSQLCommand.Parameters.Add("@ReviewStatus", SqlDbType.VarChar, 50, "ReviewStatus")
        CLMasterSQLCommand.Parameters.Add("@SampleSize", SqlDbType.Int, 4, "SampleSize")
        CLMasterSQLCommand.Parameters.Add("@ReviewedBy", SqlDbType.VarChar, 50, "ReviewedBy")
        CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 50, "SLOC")
        CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                UpOracle = False
                If CLIDs.Tables(0).Rows(i)("ReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("ReviewStatus") = ""
                If CLIDs.Tables(0).Rows(i)("OriReviewStatus") Is DBNull.Value Then CLIDs.Tables(0).Rows(i)("OriReviewStatus") = ""
                NRDate = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                OriNRDate = IIf(CLIDs.Tables(0).Rows(i)("oriNextReviewDate") Is DBNull.Value, CDate("2000-01-01"), CLIDs.Tables(0).Rows(i)("oriNextReviewDate"))

                If NRDate <> OriNRDate _
                    Or CLIDs.Tables(0).Rows(i)("ReviewStatus") <> CLIDs.Tables(0).Rows(i)("OriReviewStatus") _
                    Or CLIDs.Tables(0).Rows(i)("SampleSize") <> CLIDs.Tables(0).Rows(i)("OriSampleSize") Then

                    If CLIDs.Tables(0).Rows(i)("ReviewStatus") = "Under Review" Then

                        QtyDiff = CLIDs.Tables(0).Rows(i)("SampleSize") - CLIDs.Tables(0).Rows(i)("OriSampleSize")
                        CLIDs.Tables(0).Rows(i)("Qty") -= QtyDiff
                        CLIDs.Tables(0).Rows(i).AcceptChanges()

                    End If

                    ''''''''''''''''''''''''''
                    DR = CLIDs.Tables(1).Select("SLOC = '" & CLIDs.Tables(0).Rows(i)("SLOC") & "'" _
                                                & " and StorageBin = '" & CLIDs.Tables(0).Rows(i)("StorageBin") & "'" _
                                                & " and MaterialNo = '" & CLIDs.Tables(0).Rows(i)("MaterialNo") & "'")

                    For j = 0 To DR.Length - 1
                        CLMasterSQLCommand.Parameters("@CLID").Value = DR(j)("CLID")
                        CLMasterSQLCommand.Parameters("@Qty").Value = DR(j)("Qty") - IIf(DR(j)("SampleSize") Is DBNull.Value, 0, DR(j)("SampleSize"))

                        If CLIDs.Tables(0).Rows(i)("NextReviewDate") Is DBNull.Value Then
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = DBNull.Value
                        Else
                            CLMasterSQLCommand.Parameters("@NextReviewDate").Value = IIf(CLIDs.Tables(0).Rows(i)("NextReviewDate") < Today(), DBNull.Value, CLIDs.Tables(0).Rows(i)("NextReviewDate"))
                        End If

                        CLMasterSQLCommand.Parameters("@ReviewStatus").Value = System.DBNull.Value

                        If DR(j)("SampleSize") Is DBNull.Value Then
                            DR(j)("SampleSize") = 0
                        End If

                        CLMasterSQLCommand.Parameters("@SampleSize").Value = DR(j)("SampleSize")
                        CLMasterSQLCommand.Parameters("@ReviewedBy").Value = OracleLoginData.ToUpper

                        CLMasterSQLCommand.Parameters("@SLOC").Value = CLIDs.Tables(0).Rows(i)("SLOC")
                        CLMasterSQLCommand.Parameters("@StorageBin").Value = CLIDs.Tables(0).Rows(i)("StorageBin")

                        CLMasterSQLCommand.CommandType = CommandType.Text
                        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                        ra = CLMasterSQLCommand.ExecuteNonQuery()

                        DR(j)("Qty") = CLMasterSQLCommand.Parameters("@Qty").Value
                        DR(j).AcceptChanges()
                    Next
                End If
            Next
            myConn.Close()
        Catch ex As Exception
            eTraceUpd_Pass = ex.Message
            ErrorLogging("SLED-eTraceUpd_Pass", OracleLoginData, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function


End Class
