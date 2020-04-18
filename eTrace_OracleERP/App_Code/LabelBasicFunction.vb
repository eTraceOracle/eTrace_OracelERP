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

Public Structure LabelData
    Public RTNo As String
    Public Material As String
    Public Description As String
    Public MatlRev As String
    Public Qty As String
    Public UOM As String
    Public StatusCode As String
    Public SubInv As String
    Public Locator As String
    Public RTLot As String
    Public PurOrdNo As String
    Public MFR As String
    Public MPN As String
    Public Stemp As String
    Public MSL As String
    Public RoHS As String
    Public DCode As String
    Public LotNo As String
    Public COO As String
    Public RecDate As String
    Public ExpDate As String
    Public Units As String
    Public LabelID As String
    Public SONo As String
    Public SOLine As String
    Public DelFlag As Boolean
    Public UpdateStatus As String
    Public ItemData As DataSet
    Public RefCLID As String
    Public StorageType As String
    Public ItemText As String
    Public InvoiceNo As String
    Public VendorID As String
End Structure

'---------------mic-----------------------------

Public Structure LabelSEQ
    Public ASSY As String
    Public DJ As String
    Public JOBID As String
    Public JOBSize As String
    Public OrgCode As String
    Public Process As String
    Public Qty As String
    Public RecDate As String
    Public SEQID As String

End Structure
'---------------mic-----------------------------

Public Structure ItemData
    Public Material As String
    Public MatlDesc As String
    Public MatlRev As String
    'Public BaseUOM As String
    'Public InspFlag As String
    'Public CommCode As String
    'Public Stemp As String
    'Public MSL As String
End Structure

Public Structure ConversionResult
    Public OracleFlag As String
    Public CLIDFlag As String
    Public PrintFlag As String
    Public ErrorMsg As String
    Public SLEDFlag As String
    Public ExpDate As String
    Public TraceLevel As String
    Public CLIDs As DataSet
End Structure

Public Structure DashBoardData
    Public RTNo As String
    Public Material As String
    Public PalletID As String
    Public Status As String
    Public StatusDesc As String
    Public BerthID As String
End Structure

Public Structure CartonLabel
    Public CartonID As String
    Public Model As String
    Public Qty As String
    Public APN As String
End Structure


Public Class LabelBasicFunction
    Inherits PublicFunction
    Private POLabelFile As String = "D:\eTrace\POLabel.lab"
    Private CartonLabelFile As String = "D:\eTrace\CartonLabel.lab"

    Public Function ReadLabelIDs(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal TransactionType As String) As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim myCLIDs As DataSet = New DataSet

            Dim ErrorTable As DataTable
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim ErrorRow As Data.DataRow

            Try
                Dim Sqlstr As String
                Dim CLIDConfig As String = ""

                Dim StartStr As String
                StartStr = Microsoft.VisualBasic.Left(CLID, 1)
                Sqlstr = String.Format("exec sp_ReadLabelIDs  '{0}','{1}','{2}' ", LoginData.OrgCode, CLID, TransactionType)
                myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")
                If TransactionType = "CH09Code" Then Return myCLIDs

                myCLIDs.Tables.Add(ErrorTable)
                If myCLIDs.Tables(0).Rows.Count = 0 Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrorMsg") = "Invalid Label ID " & CLID          'Need to add message in table T_message
                    ErrorTable.Rows.Add(ErrorRow)
                    Return myCLIDs
                End If

                If TransactionType.Contains("TRANSFER") Then Return myCLIDs

                'Read CLID Config to see if allow users to Split Caron BoxID in HH function CLID Split
                If TransactionType = "SPLIT" Then
                    If StartStr = "B" Then
                        Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID009'")
                        CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr))
                        If CLIDConfig = "NO" Then
                            'Clear up label data if not allow to split CaronID
                            myCLIDs.Tables(0).Clear()

                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow("ErrorMsg") = "Not Allow to split CaronID " & CLID          'Need to add message in table T_message
                            ErrorTable.Rows.Add(ErrorRow)
                        End If
                    End If
                    Return myCLIDs
                End If


                'Read CLID Config to see if allow users to Disable / Enable CLID in HH function CLID Update
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID001'")
                CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr))

                'If Disable / Enable CLID function is not allowed in T_Config, then check user account in table field Exempt
                If CLIDConfig = "NO" Then
                    Sqlstr = String.Format("Select Exempt from T_Config with (nolock) where ConfigID = 'CLID001'")
                    Dim UserExempt As String = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                    If UserExempt <> "" Then
                        If UserExempt.Contains(LoginData.User.ToUpper) Then CLIDConfig = "YES"
                    End If
                End If
                myCLIDs.Tables(0).Rows(0)("Config1") = CLIDConfig

                'Read CLID Config to see if allow users to change CLIDQty for Disabled CLID
                If myCLIDs.Tables(0).Rows(0)("StatusCode") = "0" Then
                    CLIDConfig = ""
                    Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID003'")
                    CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                    myCLIDs.Tables(0).Rows(0)("Config2") = CLIDConfig
                End If

                If TransactionType = "PLID UPDATE" Then Return myCLIDs

                'Read the Lists of MSL value for HH function CLID Data Update   --- 01/14/2015
                If TransactionType = "CLID UPDATE" Then
                    CLIDConfig = ""
                    Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC007'")
                    CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                    myCLIDs.Tables(0).Rows(0)("MslLists") = CLIDConfig
                End If

                Return ReadLabelAML(LoginData, myCLIDs)

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-ReadLabelIDs", LoginData.User.ToUpper, "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function ReadLabelAML(ByVal LoginData As ERPLogin, ByVal myCLIDs As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim CLID As String = myCLIDs.Tables("LabelIDs").Rows(0)("CLID").ToString

            Try

                'Record Server System Date and send back to Client for ExpDate checking, as the HHClient system date may be 1/1/2007
                myCLIDs.Tables("LabelIDs").Columns.Add(New Data.DataColumn("SystemDate", System.Type.GetType("System.String")))
                myCLIDs.Tables("LabelIDs").Columns.Add(New Data.DataColumn("AMLFlag", System.Type.GetType("System.String")))
                myCLIDs.Tables("LabelIDs").Rows(0)("SystemDate") = Now.Date.ToShortDateString


                'Read AML Config to see if allow users to update AML Data in HH function CLID Update
                Dim Sqlstr As String
                Dim AMLFlag As String = ""
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID006'")
                AMLFlag = Convert.ToString(da.ExecuteScalar(Sqlstr))

                'If Update AML Data function is not allowed in T_Config, then check user account in table field Exempt
                If AMLFlag = "NO" Then
                    Sqlstr = String.Format("Select Exempt from T_Config with (nolock) where ConfigID = 'CLID006'")
                    Dim UserExempt As String = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                    If UserExempt <> "" Then
                        If UserExempt.Contains(LoginData.User.ToUpper) Then AMLFlag = "YES"
                    End If
                End If
                myCLIDs.Tables(0).Rows(0)("AMLFlag") = AMLFlag


                Dim MaterialNo As String = myCLIDs.Tables("LabelIDs").Rows(0)("MaterialNo").ToString

                Dim ds As New DataSet
                Dim ItemList(1) As String

                ItemList(0) = MaterialNo
                ds = GetAML(ItemList)
                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Return myCLIDs
                myCLIDs.Merge(ds.Tables("AMLData"))


                'Read CLID Config to see if allow eTrace data synchronize with iPro when there is any different for RoHS / Stemp / MSL / AddlData ?
                'Since iPro Satefy information was correct now, so need to synchronize iPro data to eTrace 03/26/2013
                Dim CLIDConfig As String
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID005'")
                CLIDConfig = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                If CLIDConfig = "NO" Then Return myCLIDs


                Dim ItemData As DataTable = ds.Tables("ItemData")
                If ItemData.Rows.Count = 0 Then Return myCLIDs


                Dim RoHS, Stemp, MSL, AddlData As String
                RoHS = myCLIDs.Tables("LabelIDs").Rows(0)("RoHS").ToString
                Stemp = myCLIDs.Tables("LabelIDs").Rows(0)("Stemp").ToString
                'MSL = myCLIDs.Tables("LabelIDs").Rows(0)("MSL").ToString
                AddlData = FixNull(myCLIDs.Tables("LabelIDs").Rows(0)("AddlData"))

                Dim iProFlag As String = ""
                Dim ItemFlag As Boolean = False

                If RoHS <> ItemData.Rows(0)("RoHS").ToString Then
                    iProFlag = "1"
                    ItemFlag = True
                    RoHS = ItemData.Rows(0)("RoHS").ToString
                End If

                If Stemp <> ItemData.Rows(0)("Stemp").ToString Then
                    ItemFlag = True
                    Stemp = ItemData.Rows(0)("Stemp").ToString
                End If

                'Do NOT auto synchronize with iPro for MSL in CLID Data Update function   --- 01/14/2015
                'If MSL <> ItemData.Rows(0)("MSL").ToString Then
                '    ItemFlag = True
                '    MSL = ItemData.Rows(0)("MSL").ToString
                'End If

                If AddlData <> FixNull(ItemData.Rows(0)("AddlData")) Then
                    iProFlag = "1"
                    ItemFlag = True
                    AddlData = FixNull(ItemData.Rows(0)("AddlData"))
                End If

                If ItemFlag = True Then
                    Dim r As Integer
                    'Sqlstr = String.Format("UPDATE T_CLMaster set RoHS='{0}', Stemp='{1}', MSL='{2}', AddlData='{3}' where MaterialNo='{4}'", RoHS, Stemp, MSL, AddlData, MaterialNo)
                    Sqlstr = String.Format("UPDATE T_CLMaster set RoHS='{0}', Stemp='{1}', AddlData='{2}' where MaterialNo='{3}'", RoHS, Stemp, AddlData, MaterialNo)
                    r = da.ExecuteNonQuery(Sqlstr)
                    If r > 0 Then myCLIDs.Tables("LabelIDs").Rows(0)("Config3") = iProFlag
                End If

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-ReadLabelAML", LoginData.User.ToUpper, "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                Return myCLIDs
            End Try
        End Using

    End Function

    Public Function SplitLabelIDs(ByVal LoginData As ERPLogin, ByVal LabelID As String, ByVal LabelStatus As String, ByVal Items As DataSet) As DataSet

        SplitLabelIDs = New DataSet

        Dim ErrorTable As DataTable
        Dim myDataColumn As DataColumn
        ErrorTable = New Data.DataTable("ErrorTable")
        myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
        ErrorTable.Columns.Add(myDataColumn)

        Dim ErrorRow As Data.DataRow
        SplitLabelIDs.Tables.Add(ErrorTable)

        Dim ra As Integer
        Dim i, k As Integer
        Dim CLIDArry() As String
        Dim NewCLID, QtyUnits, SplitQty As String
        Dim myCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            QtyUnits = ""
            For i = 0 To Items.Tables("MatList").Rows.Count - 1
                If Items.Tables("MatList").Rows(i)("SplitQty") Is DBNull.Value Then
                ElseIf Items.Tables("MatList").Rows(i)("SplitQty") = 0 Then
                Else
                    SplitQty = Items.Tables("MatList").Rows(i)("SplitQty").ToString & ","
                    QtyUnits = QtyUnits & SplitQty
                End If
            Next

            myConn.Open()

            myCommand = New SqlClient.SqlCommand("sp_SplitLabelIDs", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.CommandTimeout = TimeOut_M5

            myCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20)
            myCommand.Parameters("@CLID").Direction = ParameterDirection.Input
            myCommand.Parameters("@CLID").Value = LabelID

            myCommand.Parameters.Add("@QtyUnits", SqlDbType.VarChar, 8000)
            myCommand.Parameters("@QtyUnits").Direction = ParameterDirection.Input
            myCommand.Parameters("@QtyUnits").Value = QtyUnits

            myCommand.Parameters.Add("@User", SqlDbType.VarChar, 50)
            myCommand.Parameters("@User").Direction = ParameterDirection.Input
            myCommand.Parameters("@User").Value = LoginData.User

            myCommand.Parameters.Add("@LabelStatus", SqlDbType.VarChar, 50)
            myCommand.Parameters("@LabelStatus").Direction = ParameterDirection.Input
            myCommand.Parameters("@LabelStatus").Value = LabelStatus

            myCommand.Parameters.Add("@newCLID", SqlDbType.VarChar, 8000)
            myCommand.Parameters("@newCLID").Direction = ParameterDirection.Output
            myCommand.Parameters("@newCLID").Value = LabelID

            k = 0
            NewCLID = ""

            'While (k < 3 And NewCLID = "")
            '    'Threading.Thread.Sleep(60 * 1000)
            '    Try
            '        ra = myCommand.ExecuteNonQuery()
            '        NewCLID = myCommand.Parameters("@newCLID").Value
            '    Catch ex As Exception
            '        k = k + 1
            '        ErrorLogging("Call-sp_SplitLabelIDs" & Str(k), LoginData.User & " " & Str(k), "Original LabelID: " & LabelID & ", " & ex.Message & ex.Source, "E")
            '        If ex.Message.Contains("Invalid Label ID") Then k = 4
            '    End Try
            'End While


            'Only call Stored Procedure once to avoid unexpected error in FY 06/19/2012
            Try
                ra = myCommand.ExecuteNonQuery()
                NewCLID = myCommand.Parameters("@newCLID").Value
            Catch ex As Exception
                ErrorLogging("Call-sp_SplitLabelIDs", LoginData.User, "Original LabelID: " & LabelID & ", " & ex.Message & ex.Source, "E")
                If ex.Message.Contains("Invalid Label ID") Then k = 1
            End Try
            myConn.Close()

            If NewCLID = "" Then
                ErrorLogging("SplitLabelIDs-eTraceDBUpdate", LoginData.User, "Original LabelID: " & LabelID, "I")
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "Splitting failed, please contact IT."
                If k = 1 Then ErrorRow("ErrorMsg") = "Invalid Label ID " & LabelID
                ErrorTable.Rows.Add(ErrorRow)
                Return SplitLabelIDs
            End If

            Try
                k = -1
                CLIDArry = Split(NewCLID, ",")

                For i = 0 To Items.Tables("MatList").Rows.Count - 1
                    If Items.Tables("MatList").Rows(i)("SplitQty") Is DBNull.Value Then
                    ElseIf Items.Tables("MatList").Rows(i)("SplitQty") = 0 Then
                    Else
                        k = k + 1
                        If CLIDArry.Length > k Then
                            Items.Tables("MatList").Rows(i)("CLID") = CLIDArry(k)
                        End If
                    End If
                Next
            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-SplitLabelIDs" & Str(k), LoginData.User.ToUpper & " " & Str(k), "Original LabelID: " & LabelID & ", " & ex.Message & ex.Source, "E")
            End Try

            Return Items

        Catch ex As Exception
            ErrorLogging("LabelBasicFunction-SplitLabelIDs", LoginData.User.ToUpper, "Original LabelID: " & LabelID & ", " & ex.Message & ex.Source, "E")
            ErrorRow = ErrorTable.NewRow()
            ErrorRow("ErrorMsg") = "Splitting failed, please contact IT."
            ErrorTable.Rows.Add(ErrorRow)
            Return SplitLabelIDs
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function LabelDataUpdate(ByVal LoginData As ERPLogin, ByVal myCLID As LabelData) As String
        Using da As DataAccess = GetDataAccess()

            Dim sqlstr As String
            Dim StartStr, MidStr, LabelType, QMLStatus As String
            Dim Qty As Decimal = Convert.ToDecimal(myCLID.Qty)
            Dim ExpDate As Date

            Try
                StartStr = Microsoft.VisualBasic.Left(myCLID.LabelID, 1)
                MidStr = Microsoft.VisualBasic.Mid(myCLID.LabelID, 3, 1)
                If StartStr = "B" OrElse MidStr = "P" Then
                    LabelType = "PCBID"
                Else
                    LabelType = "CLID"
                End If

                If myCLID.UpdateStatus = "DISABLE" Then
                    sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Disable', StatusCode = 0, StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}', BoxID='{1}' where CLID='{2}'", LoginData.User.ToUpper, DBNull.Value, myCLID.LabelID)
                ElseIf myCLID.UpdateStatus = "ENABLE" Then
                    sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Enable', StatusCode = 1, StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}' where CLID='{1}'", LoginData.User.ToUpper, myCLID.LabelID)
                Else

                    'Filter Special Characters for MFR / MPN
                    myCLID.MFR = FilterSpecial(myCLID.MFR)
                    myCLID.MPN = FilterSpecial(myCLID.MPN)

                    'In order to avoid update error for long DateCode / LotNo, will only intercept the characters if exceeds the Max length
                    If myCLID.DCode.Length > 20 Then myCLID.DCode = Left(myCLID.DCode, 20)
                    If myCLID.LotNo.Length > 50 Then myCLID.LotNo = Left(myCLID.LotNo, 50)

                    If LabelType = "CLID" Then
                        QMLStatus = "MODIFIED" & " by " & LoginData.User & " at " & Date.Now
                        QMLStatus = Mid(SQLString(QMLStatus), 1, 50)
                        If myCLID.ExpDate Is Nothing OrElse myCLID.ExpDate Is DBNull.Value OrElse myCLID.ExpDate.ToString = "" Then
                            sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Update', StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}',DateCode='{1}', LotNo='{2}', CountryOfOrigin='{3}', QtyBaseUOM={4}, SLOC='{5}',StorageBin='{6}',MSL='{7}'  where CLID='{8}'", LoginData.User.ToUpper, myCLID.DCode, myCLID.LotNo, myCLID.COO, Qty, myCLID.SubInv, myCLID.Locator, myCLID.MSL, myCLID.LabelID)
                            If myCLID.MPN <> "" OrElse myCLID.MFR <> "" Then
                                sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Update', StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}',DateCode='{1}', LotNo='{2}', CountryOfOrigin='{3}', QtyBaseUOM={4}, SLOC='{5}',StorageBin='{6}',Manufacturer='{7}',ManufacturerPN='{8}',QMLStatus='{9}',MSL='{10}'  where CLID='{11}'", LoginData.User.ToUpper, myCLID.DCode, myCLID.LotNo, myCLID.COO, Qty, myCLID.SubInv, myCLID.Locator, myCLID.MFR, myCLID.MPN, QMLStatus, myCLID.MSL, myCLID.LabelID)
                            End If
                        Else
                            ExpDate = CDate(myCLID.ExpDate)
                            sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Update', StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}',DateCode='{1}', LotNo='{2}', CountryOfOrigin='{3}', QtyBaseUOM={4}, SLOC='{5}',StorageBin='{6}',ExpDate='{7}',MSL='{8}' where CLID='{9}'", LoginData.User.ToUpper, myCLID.DCode, myCLID.LotNo, myCLID.COO, Qty, myCLID.SubInv, myCLID.Locator, myCLID.ExpDate, myCLID.MSL, myCLID.LabelID)
                            If myCLID.MPN <> "" OrElse myCLID.MFR <> "" Then
                                sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Update', StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}',DateCode='{1}', LotNo='{2}', CountryOfOrigin='{3}', QtyBaseUOM={4}, SLOC='{5}',StorageBin='{6}',ExpDate='{7}', Manufacturer='{8}',ManufacturerPN='{9}',QMLStatus='{10}',MSL='{11}' where CLID='{12}'", LoginData.User.ToUpper, myCLID.DCode, myCLID.LotNo, myCLID.COO, Qty, myCLID.SubInv, myCLID.Locator, myCLID.ExpDate, myCLID.MFR, myCLID.MPN, QMLStatus, myCLID.MSL, myCLID.LabelID)
                            End If
                        End If
                    Else
                        sqlstr = String.Format("UPDATE T_CLMaster set LastTransaction='CLID Update', StorageType=NULL, ChangedOn=getDate(),ChangedBy='{0}',DateCode='{1}', LotNo='{2}', QtyBaseUOM={3}, SLOC='{4}',StorageBin='{5}', SONo='{6}', SOLine='{7}' where CLID='{8}'", LoginData.User.ToUpper, myCLID.DCode, myCLID.LotNo, Qty, myCLID.SubInv, myCLID.Locator, myCLID.SONo, myCLID.SOLine, myCLID.LabelID)
                    End If

                End If

                If LabelType = "PCBID" AndAlso myCLID.UpdateStatus = "UPDATE" Then
                    Dim SONo, SOLine, DefSO, DefLine As String
                    SONo = myCLID.SONo.ToString
                    SOLine = myCLID.SOLine.ToString
                    'DefSO = myCLID.DCode.ToString                 'Default SONo
                    'DefLine = myCLID.LotNo.ToString               'Default SOLine
                    DefSO = myCLID.RTNo.ToString                   'Default RTNo
                    DefLine = myCLID.RTLot.ToString                'Default RTLot

                    If SONo <> "" And (SONo <> DefSO OrElse (SOLine <> "" AndAlso SOLine <> DefLine)) Then
                        Dim dsSO As DataSet = New DataSet
                        dsSO = GetSOLine(LoginData, SONo, SOLine)

                        If dsSO Is Nothing OrElse dsSO.Tables.Count = 0 OrElse dsSO.Tables(0).Rows.Count = 0 Then
                            LabelDataUpdate = "Invalid SO Number or SO Line"
                            Exit Function
                        End If

                        If dsSO.Tables(0).Rows(0)("o_error_mssg").ToString <> "" Then
                            LabelDataUpdate = dsSO.Tables(0).Rows(0)("o_error_mssg").ToString
                            Exit Function
                        End If

                        Dim DR() As DataRow = Nothing
                        DR = dsSO.Tables(0).Select("ordered_item = '" & myCLID.Material.ToUpper & "'")
                        If DR.Length = 0 Then
                            LabelDataUpdate = "Invalid SO Number or SO Line for item: " & myCLID.Material.ToUpper
                            Exit Function
                        End If
                        'If DR(0)("ordered_quantity") < Qty Then
                        '    LabelDataUpdate = "Current Qty should not exceed the rest SOLine Qty: " & DR(0)("ordered_quantity").ToString
                        '    Exit Function
                        'End If

                        'Don't know if there has multiple SOLine for the same item when user didn't input SOLine, so need to check the Total Order Qty here            5/8/2014
                        'Dim i As Integer
                        'Dim OrderQty As Decimal = 0
                        'For i = 0 To DR.Length - 1
                        '    'If OrderQty < DR(i)("ordered_quantity") Then OrderQty = DR(i)("ordered_quantity")
                        '    OrderQty = OrderQty + DR(i)("ordered_quantity")
                        'Next
                        'If OrderQty < Qty Then
                        '    Dim ErrMsg As String = "Current Qty should not exceed the rest SOLine Qty: "
                        '    If SOLine = "" Then ErrMsg = "Current Qty should not exceed the rest SO Qty: "
                        '    LabelDataUpdate = ErrMsg & OrderQty.ToString
                        '    Exit Function
                        'End If

                    End If

                End If

                da.ExecuteNonQuery(sqlstr)
                LabelDataUpdate = "Y"

                'Record CLID Data Update into table T_PO_CLID_RTN
                If LabelType = "CLID" Then
                    Dim RtnMsg As String = ""
                    sqlstr = String.Format("exec sp_RecordCLIDUpdate '{0}','{1}' ", myCLID.LabelID, LoginData.User.ToUpper)
                    RtnMsg = da.ExecuteScalar(sqlstr).ToString
                End If

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-LabelDataUpdate", LoginData.User.ToUpper, "CLID: " & myCLID.LabelID & ", " & ex.Message & ex.Source, "E")
                LabelDataUpdate = "Update error"
            End Try

        End Using

    End Function

    Public Function ReadPalletID(ByVal LoginData As ERPLogin, ByVal LabelID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim myPalletID As DataSet = New DataSet

            Dim ErrorTable As DataTable
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim LabelType As String
            Dim ErrorRow As Data.DataRow

            Try
                If Left(LabelID, 1) = "P" Then
                    LabelType = "BoxID"
                Else
                    LabelType = "CLID"
                End If

                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_ReadPalletID  '{0}', '{1}', '{2}', '{3}'", LoginData.OrgCode, LabelType, LabelID, "1")   'Read valid BoxID
                myPalletID = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                myPalletID.Tables.Add(ErrorTable)
                If myPalletID.Tables(0).Rows.Count > 0 Then Return myPalletID

                If LabelType = "CLID" Then
                    If myPalletID.Tables(0).Rows.Count = 0 Then
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow("ErrorMsg") = "Invalid BoxID " & LabelID
                        ErrorTable.Rows.Add(ErrorRow)
                    End If
                    Return myPalletID
                End If

                If myPalletID.Tables(0).Rows.Count = 0 Then
                    Dim ds As New DataSet
                    Sqlstr = String.Format("exec sp_ReadPalletID  '{0}', '{1}', '{2}', '{3}'", LoginData.OrgCode, LabelType, LabelID, "0")  'Read invalid BoxID
                    ds = da.ExecuteDataSet(Sqlstr, "LabelIDs")
                    If ds.Tables(0).Rows.Count > 0 Then
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow("ErrorMsg") = "PalletID " & LabelID & " is inactive."
                        ErrorTable.Rows.Add(ErrorRow)
                        Return myPalletID
                    End If

                    ds = New DataSet
                    Sqlstr = String.Format("exec sp_CheckPalletID  '{0}', '{1}'", LoginData.OrgCode, LabelID)  'Check if PalletID exists in T_Shipment
                    ds = da.ExecuteDataSet(Sqlstr, "LabelIDs")
                    If ds.Tables(0).Rows.Count > 0 Then
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow("ErrorMsg") = "PalletID " & LabelID & " is used in Production now."
                        ErrorTable.Rows.Add(ErrorRow)
                        Return myPalletID
                    End If

                End If

                Return myPalletID

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-ReadPalletID", LoginData.User.ToUpper, "LabelID: " & LabelID & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function PalletManagement(ByVal LoginData As ERPLogin, ByVal myCLIDs As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Dim PalletID As String = myCLIDs.Tables(0).Rows(0)("PalletID").ToString

            Try
                Dim Sqlstr As String
                myCLIDs.DataSetName = "Items"

                Sqlstr = String.Format("exec sp_PalletManagement '{0}','{1}',N'{2}'", LoginData.OrgCode, LoginData.User.ToUpper, DStoXML(myCLIDs))
                PalletManagement = da.ExecuteScalar(Sqlstr).ToString
            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-PalletManagement", LoginData.User.ToUpper, "PalletID: " & PalletID & ", " & ex.Message & ex.Source, "E")
                PalletManagement = "Data Update error for PalletID " & PalletID
            End Try

        End Using

    End Function

    Public Function GetCartonLabel(ByVal LoginData As ERPLogin, ByVal CartonID As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            GetCartonLabel = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetCartonLabel  '{0}','{1}' ", LoginData.OrgCode, CartonID)
                GetCartonLabel = da.ExecuteDataSet(Sqlstr, "LabelIDs")
                Return GetCartonLabel

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-GetCartonLabel", LoginData.User.ToUpper, "CartonID: " & CartonID & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function PrintCartonLabel(ByVal myCLID As DataSet, ByVal Printer As String) As Boolean
        Using da As DataAccess = GetDataAccess()

            Dim CartonID As String = myCLID.Tables(0).Rows(0)("CartonID").ToString

            Try

                Dim CustomerLabel As CartonLabel = New CartonLabel

                CustomerLabel.CartonID = CartonID
                CustomerLabel.Model = myCLID.Tables(0).Rows(0)("Model").ToString
                CustomerLabel.Qty = myCLID.Tables(0).Rows(0)("Qty").ToString
                CustomerLabel.APN = myCLID.Tables(0).Rows(0)("APN").ToString

                Dim sql As String
                Dim strContent As String
                Dim arryFile() As String
                Try
                    arryFile = Split(CartonLabelFile, "\")
                    strContent = "CartonID^" & CustomerLabel.CartonID & "^Model^" & CustomerLabel.Model & "^Qty^" & CustomerLabel.Qty & "^APN^" & CustomerLabel.APN

                    sql = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", arryFile(UBound(arryFile)), Printer, SQLString(strContent))
                    da.ExecuteScalar(sql)
                    PrintCartonLabel = True
                Catch ex As Exception
                    ErrorLogging("LabelBasicFunction-PrintCartonLabel1", "", "CartonID: " & CartonID & ", " & ex.Message & ex.Source, "E")
                    PrintCartonLabel = False
                End Try

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-PrintCartonLabel", "", "CartonID: " & CartonID & ", " & ex.Message & ex.Source, "E")
                PrintCartonLabel = False
            End Try

        End Using

    End Function

    Public Function GetItemDescription(ByVal LoginData As ERPLogin, ByVal PartNo As String) As ItemData
        Using da As DataAccess = GetDataAccess()
            GetItemDescription.Material = PartNo
            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_wip_pkg.get_item_desc"
                OC.Parameters.Add("p_org_code", OracleType.VarChar).Value = LoginData.OrgCode
                OC.Parameters.Add("p_item_num", OracleType.VarChar).Value = PartNo.ToUpper()
                OC.Parameters.Add("o_desc", OracleType.VarChar, 500)
                OC.Parameters.Add("o_rev", OracleType.VarChar, 20)
                OC.Parameters("o_desc").Direction = ParameterDirection.Output
                OC.Parameters("o_rev").Direction = ParameterDirection.Output
                If OC.Connection.State = ConnectionState.Closed Then
                    OC.Connection.Open()
                End If
                aa = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                GetItemDescription.MatlDesc = OC.Parameters("o_desc").Value.ToString
                GetItemDescription.MatlRev = OC.Parameters("o_rev").Value.ToString

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-GetItemDescription", "", "PartNo: " & PartNo & ", " & oe.Message & oe.Source, "E")
                GetItemDescription.MatlDesc = ""
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try

        End Using
    End Function

    Public Function GetItemMaster(ByVal LoginData As ERPLogin, ByVal PartNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("item_data")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_wip_pkg.get_item_master"              'Get Item Master Data
                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_item_num", OracleType.VarChar, 240).Value = PartNo.ToUpper()

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "item_data")
                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-GetItemMaster", "", "PartNo: " & PartNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function PartNoTransfer(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As DataSet = New DataSet
            Dim myCLIDs As DataSet = New DataSet

            Dim PNData As New DataTable
            PNData = Items.Tables(0).Clone()
            ds.Tables.Add(PNData)

            Dim CLIDS As DataTable
            Dim dsItem As DataSet = New DataSet

            CLIDS = New Data.DataTable("CLIDS")
            CLIDS.Columns.Add(New Data.DataColumn("CLID", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("NewPartNo", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("NewPNDesc", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("DestSubInv", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("DestLocator", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("StorageType", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("BoxID", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("LotCtrlCode", System.Type.GetType("System.String")))
            CLIDS.Columns.Add(New Data.DataColumn("MCPosition", System.Type.GetType("System.String")))
            dsItem.Tables.Add(CLIDS)

            Dim i, j, k As Integer
            Dim Sqlstr, RTLot As String
            Dim DR() As DataRow = Nothing
            Dim ErrorMsg As String = "Error for Item Transfer: from " & Items.Tables(0).Rows(0)("MaterialNo") & " to " & Items.Tables(0).Rows(0)("NewPartNo")

            Dim ErrorRow As Data.DataRow
            Dim ErrorTable As New DataTable
            ErrorTable.Columns.Add(New Data.DataColumn("ErrMsg", System.Type.GetType("System.String")))
            PartNoTransfer = New DataSet
            PartNoTransfer.Tables.Add(ErrorTable)

            Try

                j = -1
                For i = 0 To Items.Tables(0).Rows.Count - 1
                    'Save CLID / BoxID Data for Store Procedure update later
                    CLIDS.ImportRow(Items.Tables(0).Rows(i))

                    RTLot = Items.Tables(0).Rows(i)("RTLot").ToString
                    If Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                        Sqlstr = " RTLot = '" & RTLot & "'"
                    Else
                        Sqlstr = " RTLot = '" & RTLot & "' and ExpDate = '" & Items.Tables(0).Rows(i)("ExpDate") & "'"
                    End If

                    DR = ds.Tables(0).Select(Sqlstr)
                    If DR.Length = 0 Then
                        j = j + 1
                        ds.Tables(0).ImportRow(Items.Tables(0).Rows(i))
                        ds.Tables(0).Rows(j)("CLID") = ""
                        ds.Tables(0).Rows(j)("BoxID") = ""
                        ds.Tables(0).Rows(j).AcceptChanges()
                        ds.Tables(0).Rows(j).SetAdded()
                    Else
                        Dim Qty As Decimal = DR(0)("QtyBaseUOM")
                        DR(0)("QtyBaseUOM") = Qty + Items.Tables(0).Rows(i)("QtyBaseUOM")
                        DR(0).AcceptChanges()
                        DR(0).SetAdded()
                    End If
                Next

                Dim OraTransfer As String = ""
                Try
                    OraTransfer = Oracle_PNTransfer(LoginData, ds)
                Catch ex As Exception
                    ErrorLogging("PartNoTransfer-Call-Oracle_PNTransfer", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                    OraTransfer = ex.Message
                End Try
                If OraTransfer <> "Y" Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrMsg") = OraTransfer
                    ErrorTable.Rows.Add(ErrorRow)
                    Return PartNoTransfer
                End If

                Try
                    dsItem.DataSetName = "Items"
                    Sqlstr = String.Format("exec sp_UpdateItemTransfer '{0}','{1}',N'{2}'", LoginData.OrgCode, LoginData.User.ToUpper, DStoXML(dsItem))
                    myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDS")
                Catch ex As Exception
                    ErrorLogging("PartNoTransfer-sp_UpdateItemTransfer", LoginData.User, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                    Return Nothing
                End Try

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-PartNoTransfer", LoginData.User.ToUpper, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function Oracle_PNTransfer(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As OracleString
            Dim oda As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Dim ErrorMsg As String = "Item Transfer from " & p_ds.Tables(0).Rows(0)("MaterialNo") & " to " & p_ds.Tables(0).Rows(0)("NewPartNo")

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                comm.ExecuteOracleNonQuery(aa)
                comm.Parameters.Clear()

                comm.CommandText = "apps.xxetr_wip_pkg.pn_transfer"
                comm.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgID  'LoginData.OrgCode
                comm.Parameters.Add("p_item_num_a", OracleType.VarChar, 50)
                comm.Parameters.Add("p_item_rev_a", OracleType.VarChar, 50)
                comm.Parameters.Add("p_lot_num_a", OracleType.VarChar, 50)
                comm.Parameters.Add("p_quantity", OracleType.Double)
                comm.Parameters.Add("p_uom", OracleType.VarChar, 50)
                comm.Parameters.Add("p_sub_a", OracleType.VarChar, 50)
                comm.Parameters.Add("p_loc_a", OracleType.VarChar, 50)
                comm.Parameters.Add("p_item_num_b", OracleType.VarChar, 50)
                comm.Parameters.Add("p_item_rev_b", OracleType.VarChar, 50)
                comm.Parameters.Add("p_sub_b", OracleType.VarChar, 50)
                comm.Parameters.Add("p_loc_b", OracleType.VarChar, 50)
                comm.Parameters.Add("p_exp_date", OracleType.VarChar, 50)
                comm.Parameters.Add("p_reference_a", OracleType.VarChar, 240)                           'Add Reference A/B
                comm.Parameters.Add("p_reference_b", OracleType.VarChar, 240)

                'comm.Parameters.Add("p_miscissue_accname", OracleType.VarChar, 240).Value = "EXCESS ISSUE TO PROD"
                'comm.Parameters.Add("p_miscrcpt_accname", OracleType.VarChar, 240).Value = "EXCESS RECPT FM PROD"
                'comm.Parameters.Add("p_miscissue_accname", OracleType.VarChar, 240).Value = "PN CONVERSION(ISSUE)"
                'comm.Parameters.Add("p_miscrcpt_accname", OracleType.VarChar, 240).Value = "PN CONVERSION(RCPT)"
                ''comm.Parameters.Add("p_miscissue_accname", OracleType.VarChar, 240).Value = "INTRNL PN SWAP ISSUE"
                ''comm.Parameters.Add("p_miscrcpt_accname", OracleType.VarChar, 240).Value = "INTRNL PN SWAP RCPT"

                ''''//-----------------------------------------------------------------------

                ''''//  The Handhelp program "FrmItemTransfer" will invoke it

                ''''//  <copyright>Copyright (c) EMRSN. All rights reserved.</copyright>

                ''''//  <date>03-28-2011</date>

                ''''//  <author>Jackson Huang</author>

                ''''//  <comment>

                ''''//  Modify("p_miscissue_accname" And "p_miscrcpt_accname") 's value 

                ''''//  by hard code

                ''''//  </comment>

                ''''//-----------------------------------------------------------------------

                comm.Parameters.Add("p_miscrcpt_accname", OracleType.VarChar, 240).Value = "PN CONVERSION(RCPT)"
                comm.Parameters.Add("p_miscissue_accname", OracleType.VarChar, 240).Value = "PN CONVERSION(ISSUE)"

                comm.Parameters.Add("o_success_flag", OracleType.VarChar, 50)
                comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 500)

                comm.Parameters("o_success_flag").Direction = ParameterDirection.Output
                comm.Parameters("o_error_mssg").Direction = ParameterDirection.Output

                comm.Parameters("p_item_num_a").SourceColumn = "MaterialNo"
                comm.Parameters("p_item_rev_a").SourceColumn = "MaterialRevision"
                comm.Parameters("p_lot_num_a").SourceColumn = "RTLot"
                comm.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
                comm.Parameters("p_uom").SourceColumn = "BaseUOM"
                comm.Parameters("p_sub_a").SourceColumn = "SLOC"
                comm.Parameters("p_loc_a").SourceColumn = "StorageBin"
                comm.Parameters("p_item_num_b").SourceColumn = "NewPartNo"
                comm.Parameters("p_item_rev_b").SourceColumn = "NewRev"
                comm.Parameters("p_sub_b").SourceColumn = "DestSubInv"
                comm.Parameters("p_loc_b").SourceColumn = "DestLocator"
                comm.Parameters("p_exp_date").SourceColumn = "ExpDate"
                comm.Parameters("p_reference_a").SourceColumn = "ReferenceA"                            'Add Reference A/B
                comm.Parameters("p_reference_b").SourceColumn = "ReferenceB"
                comm.Parameters("o_success_flag").SourceColumn = "Status"
                comm.Parameters("o_error_mssg").SourceColumn = "Message"

                oda.InsertCommand = comm
                oda.Update(p_ds.Tables(0))

                Dim Flag, Msg As String
                Flag = comm.Parameters("o_success_flag").Value.ToString
                Msg = comm.Parameters("o_error_mssg").Value.ToString

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables(0).Select("Status = 'N' or Status = ' ' or Status IS NULL ")
                If DR.Length = 0 Then
                    comm.Transaction.Commit()
                    comm.Connection.Close()
                    Return DirectCast(Flag, String)
                    Exit Function
                End If

                'Record error message to eTrace database
                Dim i As Integer
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString.Trim <> "" Then
                        If Msg = "" Then Msg = DR(i)("Message").ToString.Trim
                        Dim ErrMsg As String = "RTLot: " & DR(i)("RTLot").ToString & " with Item " & DR(i)("MaterialNo") & " and return message: "
                        ErrMsg = ErrMsg & DR(i)("Message").ToString.Trim
                        ErrorLogging("LabelBasicFunction-Oracle_PNTransfer1", LoginData.User.ToUpper, ErrorMsg & ", " & ErrMsg, "I")
                    End If
                Next
                Return Msg

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Oracle_PNTransfer", LoginData.User.ToUpper, ErrorMsg & " with exception: " & oe.Message & oe.Source, "E")
                comm.Transaction.Rollback()
                Return "Part Number Transfer error"
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
        End Using

    End Function

    Public Function InterOrgTransfer(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            InterOrgTransfer = ""

            Dim ds As DataSet = New DataSet
            Dim CLIDS As DataTable = Items.Tables(0).Clone()
            ds.Tables.Add(CLIDS)

            Dim i As Integer
            Dim RTLot As String
            Dim DR() As DataRow
            Dim myDataRow As Data.DataRow
            Dim ErrorMsg As String = "Error for InterOrg Transfer with Item " & Items.Tables("CLIDS").Rows(0)("MatNo") & " from Org " & LoginData.OrgCode & " To DestOrg " & Items.Tables("CLIDS").Rows(0)("DestOrg")

            Try

                For i = 0 To Items.Tables("CLIDS").Rows.Count - 1
                    'Make sure the RowState is SetAdded
                    If Items.Tables("CLIDS").Rows(i).RowState = DataRowState.Unchanged Then
                        Items.Tables("CLIDS").Rows(i).SetAdded()
                    End If

                    RTLot = Items.Tables("CLIDS").Rows(i)("RTLot").ToString

                    DR = Nothing
                    DR = ds.Tables("CLIDS").Select(" RTLot = '" & RTLot & "'")
                    If DR.Length = 0 Then
                        myDataRow = ds.Tables("CLIDS").NewRow()
                        myDataRow("DestOrg") = Items.Tables("CLIDS").Rows(i)("DestOrg")
                        myDataRow("CLID") = ""
                        myDataRow("MatNo") = Items.Tables("CLIDS").Rows(i)("MatNo")
                        myDataRow("QtyBaseUOM") = Items.Tables("CLIDS").Rows(i)("QtyBaseUOM")
                        myDataRow("BaseUOM") = Items.Tables("CLIDS").Rows(i)("BaseUOM")
                        myDataRow("Rev") = Items.Tables("CLIDS").Rows(i)("Rev")
                        myDataRow("RTLot") = Items.Tables("CLIDS").Rows(i)("RTLot")
                        myDataRow("SubInv") = Items.Tables("CLIDS").Rows(i)("SubInv")
                        myDataRow("Locator") = Items.Tables("CLIDS").Rows(i)("Locator")
                        myDataRow("DestSubInv") = Items.Tables("CLIDS").Rows(i)("DestSubInv")
                        myDataRow("DestLocator") = Items.Tables("CLIDS").Rows(i)("DestLocator")
                        myDataRow("LotCtrlCode") = 0
                        myDataRow("ReasonCode") = Items.Tables("CLIDS").Rows(i)("ReasonCode")
                        myDataRow("Reference") = Items.Tables("CLIDS").Rows(i)("Reference")
                        myDataRow("Status") = ""
                        myDataRow("Message") = ""

                        ds.Tables("CLIDS").Rows.Add(myDataRow)
                    Else
                        Dim UnitQty As Decimal
                        UnitQty = DR(0)("QtyBaseUOM")

                        DR(0)("QtyBaseUOM") = UnitQty + Items.Tables("CLIDS").Rows(i)("QtyBaseUOM")
                        DR(0).AcceptChanges()
                        DR(0).SetAdded()

                    End If
                Next

                Dim OraTransfer As DataSet = New DataSet
                Dim UserID As String = LoginData.User.ToUpper

                Try
                    OraTransfer = Oracle_IntOrgTransfer(LoginData, ds)
                Catch ex As Exception
                    ErrorLogging("InterOrgTransfer-Call-Oracle_IntOrgTransfer", LoginData.User.ToUpper, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                End Try

                DR = OraTransfer.Tables("CLIDS").Select("Status = 'N' or Status = ' ' or Status IS NULL ")
                If DR.Length > 0 Then
                    InterOrgTransfer = DR(0)("Message").ToString
                    If InterOrgTransfer = "" Then InterOrgTransfer = "No data is submitted in Oracle"
                    Return InterOrgTransfer
                End If

                Dim LotCtrlCode As Integer   'LotCtrlCode = "1": No LotConrol; LotCtrlCode = "2": FullConrol;
                LotCtrlCode = CInt(OraTransfer.Tables("CLIDS").Rows(0)("LotCtrlCode"))

                Dim ConfigOrg As String
                Dim DestOrg As String = Items.Tables(0).Rows(0)("DestOrg").ToString

                Dim Sqlstr As String
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'ORGLIST' ")
                ConfigOrg = Convert.ToString(da.ExecuteScalar(Sqlstr))

                Dim LocalTransfer As Integer = 0
                Dim OrgIndex As Integer = ConfigOrg.IndexOf(DestOrg)        ' If not found, OrgIndex = - 1
                If ConfigOrg <> "ALL" AndAlso OrgIndex < 0 Then
                    LocalTransfer = 1
                End If

                Dim sda As SqlClient.SqlDataAdapter = da.Sda_Insert()

                Try
                    sda.InsertCommand.CommandType = CommandType.StoredProcedure
                    sda.InsertCommand.CommandText = "sp_UpdateInterOrgTransfer"
                    sda.InsertCommand.CommandTimeout = TimeOut_M5

                    sda.InsertCommand.Parameters.Add("@DestOrg", SqlDbType.VarChar, 50).Value = DestOrg
                    sda.InsertCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20)
                    sda.InsertCommand.Parameters.Add("@SubInv", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@Locator", SqlDbType.VarChar, 20)
                    sda.InsertCommand.Parameters.Add("@StorageType", SqlDbType.VarChar, 20)
                    sda.InsertCommand.Parameters.Add("@CLIDQty", SqlDbType.Decimal)
                    sda.InsertCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50)
                    sda.InsertCommand.Parameters.Add("@LotCtrlCode", SqlDbType.Int).Value = CInt(LotCtrlCode)
                    sda.InsertCommand.Parameters.Add("@User", SqlDbType.VarChar, 50).Value = LoginData.User.ToUpper
                    sda.InsertCommand.Parameters.Add("@LocalTransfer", SqlDbType.Int).Value = CInt(LocalTransfer)
                    sda.InsertCommand.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 2000)

                    sda.InsertCommand.Parameters("@CLID").SourceColumn = "CLID"
                    sda.InsertCommand.Parameters("@SubInv").SourceColumn = "DestSubInv"
                    sda.InsertCommand.Parameters("@Locator").SourceColumn = "DestLocator"
                    sda.InsertCommand.Parameters("@StorageType").SourceColumn = "StorageType"
                    sda.InsertCommand.Parameters("@CLIDQty").SourceColumn = "QtyBaseUOM"
                    sda.InsertCommand.Parameters("@ReasonCode").SourceColumn = "ReasonCode"
                    sda.InsertCommand.Parameters("@ReturnMsg").SourceColumn = "Message"
                    sda.InsertCommand.Parameters("@ReturnMsg").Direction = ParameterDirection.Output

                    sda.InsertCommand.Connection.Open()
                    sda.Update(Items.Tables(0))
                    sda.InsertCommand.Connection.Close()

                Catch ex As Exception
                    ErrorLogging("IntOrgTransfer-sp_UpdateInterOrgTransfer", LoginData.User, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                    InterOrgTransfer = "InterOrg Transfer error"
                Finally
                    If sda.InsertCommand.Connection.State <> ConnectionState.Closed Then sda.InsertCommand.Connection.Close()
                End Try
                If InterOrgTransfer <> "" Then Return InterOrgTransfer


                DR = Items.Tables(0).Select("Message = 'N' or Message = ' ' or Message IS NULL ")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        InterOrgTransfer = DR(i)("Message").ToString
                        If InterOrgTransfer <> "" Then
                            ErrorLogging("LabelBasicFunction-InterOrgTransfer1", LoginData.User.ToUpper, ErrorMsg & " with error message: " & InterOrgTransfer, "I")
                            Exit For
                        End If
                    Next
                    If InterOrgTransfer = "" Then InterOrgTransfer = "InterOrg Transfer error"
                    Return InterOrgTransfer
                    Exit Function
                End If

                'Set Flag: Y if successful
                'InterOrgTransfer = "Y"

                'Return LotControlCode if successful.  
                'LotCtrlCode = "1": No LotConrol; LotCtrlCode = "2": FullConrol;
                InterOrgTransfer = LotCtrlCode

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-InterOrgTransfer", LoginData.User.ToUpper, ErrorMsg & "; " & ex.Message & ex.Source, "E")
                InterOrgTransfer = "InterOrg Transfer error"
            End Try
        End Using

    End Function

    Public Function Oracle_IntOrgTransfer(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim aa As OracleString
            Dim oda As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Dim ErrorMsg As String = "Error for InterOrg Transfer with Item " & p_ds.Tables("CLIDS").Rows(0)("MatNo") & " from Org " & LoginData.OrgCode & " To DestOrg " & p_ds.Tables("CLIDS").Rows(0)("DestOrg")

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim DestOrg As String = p_ds.Tables("CLIDS").Rows(0)("DestOrg").ToString
                Dim DestOrgID As String = GetOrgID(DestOrg)
                comm.CommandType = CommandType.StoredProcedure
                comm.CommandText = "apps.xxetr_wip_pkg.initialize"
                comm.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                comm.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                comm.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                comm.ExecuteOracleNonQuery(aa)
                comm.Parameters.Clear()

                comm.CommandText = "apps.xxetr_wip_pkg.org_transfer"
                comm.Parameters.Add("p_item_num", OracleType.VarChar, 240)
                comm.Parameters.Add("p_item_rev", OracleType.VarChar, 240)
                comm.Parameters.Add("p_quantity", OracleType.Double)
                comm.Parameters.Add("p_uom", OracleType.VarChar, 240)
                comm.Parameters.Add("p_lotnum", OracleType.VarChar, 240)
                comm.Parameters.Add("p_reason", OracleType.VarChar, 240)
                comm.Parameters.Add("p_org_code_fm", OracleType.VarChar, 240).Value = OrgID  'LoginData.OrgCode
                comm.Parameters.Add("p_subinv_fm", OracleType.VarChar, 240)
                comm.Parameters.Add("p_locator_fm", OracleType.VarChar, 240)
                comm.Parameters.Add("p_org_code_to", OracleType.VarChar, 240).Value = DestOrgID
                comm.Parameters.Add("p_subinv_to", OracleType.VarChar, 240)
                comm.Parameters.Add("p_locator_to", OracleType.VarChar, 240)
                comm.Parameters.Add("p_reference", OracleType.VarChar, 240)
                comm.Parameters.Add("o_lot_ctrl_code", OracleType.Int32)
                comm.Parameters.Add("o_success_flag", OracleType.VarChar, 240)
                comm.Parameters.Add("o_error_mssg", OracleType.VarChar, 240)

                comm.Parameters("o_lot_ctrl_code").Direction = ParameterDirection.InputOutput
                comm.Parameters("o_success_flag").Direction = ParameterDirection.InputOutput
                comm.Parameters("o_error_mssg").Direction = ParameterDirection.InputOutput

                comm.Parameters("p_item_num").SourceColumn = "MatNo"
                comm.Parameters("p_item_rev").SourceColumn = "Rev"
                comm.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
                comm.Parameters("p_uom").SourceColumn = "BaseUOM"
                comm.Parameters("p_lotnum").SourceColumn = "RTLot"
                comm.Parameters("p_reason").SourceColumn = "ReasonCode"
                comm.Parameters("p_subinv_fm").SourceColumn = "SubInv"
                comm.Parameters("p_locator_fm").SourceColumn = "Locator"
                'comm.Parameters("p_org_code_to").SourceColumn = "DestOrg"
                comm.Parameters("p_subinv_to").SourceColumn = "DestSubInv"
                comm.Parameters("p_locator_to").SourceColumn = "DestLocator"
                comm.Parameters("p_reference").SourceColumn = "Reference"
                comm.Parameters("o_lot_ctrl_code").SourceColumn = "LotCtrlCode"
                comm.Parameters("o_success_flag").SourceColumn = "Status"
                comm.Parameters("o_error_mssg").SourceColumn = "Message"

                oda.InsertCommand = comm
                oda.Update(p_ds.Tables("CLIDS"))

                Dim Flag, Msg As String
                Flag = comm.Parameters("o_success_flag").Value.ToString
                Msg = comm.Parameters("o_error_mssg").Value.ToString

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables("CLIDS").Select("Status = 'N' or Status = ' ' or Status IS NULL ")
                If DR.Length = 0 Then
                    comm.Transaction.Commit()
                    comm.Connection.Close()
                Else
                    ErrorLogging("LabelBasicFunction-Oracle_IntOrgTransfer", LoginData.User.ToUpper, ErrorMsg & " with error message: " & Msg, "I")
                    comm.Transaction.Rollback()
                    comm.Connection.Close()
                End If

                Return p_ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Oracle_IntOrgTransfer", LoginData.User.ToUpper, ErrorMsg & " with exception: " & oe.Message & oe.Source, "E")
                comm.Transaction.Rollback()
                Return p_ds
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
        End Using

    End Function

    Public Function LabelConversion(ByVal LoginData As ERPLogin, ByVal LabelPrinter As String, ByVal PrintLabels As Boolean, ByVal LabelInfo As LabelData) As ConversionResult

        Dim dsAML = New DataSet
        Dim AMLFlag As Boolean = False
        Dim RoHSFlag As Boolean = False

        'For Raw Item, check if Item Onhand Qty >= Total generation Qty ( =Qty * NoOfPackage) in the SubInv / Locator for the input RTLot, if not, return error message
        'For SubAssy, check if Item Onhand Qty >= Total generation Qty ( =Qty * NoOfPackage) in the SubInv / Locator, if not, return error message
        'For FG, check if Item/Revision Onhand Qty >= Total generation Qty ( =Qty * NoOfPackage) in the SubInv / Locator, if not, return error message

        'LabelConversion.ErrorMsg = "Item OnHand Qty is not enough "
        'If Not LabelConversion.ErrorMsg Is Nothing And LabelConversion.ErrorMsg <> "" Then Exit Function

        'ErrorLogging("IN ExpDate", LoginData.User.ToUpper, LabelInfo.ExpDate.ToString, "I")

        If LabelInfo.ItemData.Tables(0).Rows(0)("type_name") = "RM" Then
            Dim ItemList(1) As String
            ItemList(0) = LabelInfo.Material
            Try
                dsAML = GetAML(ItemList)
            Catch ex As Exception
                dsAML = Nothing
            End Try

            If dsAML Is Nothing OrElse dsAML.Tables.Count = 0 Then
            ElseIf dsAML.Tables.Count > 0 Then
                If dsAML.Tables("AMLData").Rows.Count > 0 Then AMLFlag = True
                If dsAML.Tables("ItemData").Rows.Count > 0 Then RoHSFlag = True
            End If
        End If

        Dim lblPrint As String
        If PrintLabels = True Then     ' To ensure timely feedback, Do label printing later
            lblPrint = "True"
        Else
            lblPrint = "Disabled"
        End If

        If LabelInfo.Units Is Nothing Then
            LabelInfo.Units = "1"
        End If


        Dim CLIDs As DataSet                     'Set CLIDs DataSet to save table
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
        If LabelInfo.ExpDate <> "N" Then
            NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RTLot,CreatedBy,CreatedOn,RecDate,Printed,ExpDate,ProdDate,RoHS,ReasonCode, StockType,MaterialDesc,SLOC,StorageBin,Operator,IsTraceable,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RTLot,@CreatedBy,getdate(),getdate(),@Printed,@ExpDate,@ProdDate,@RoHS,@ReasonCode,@StockType,@MaterialDesc,@SLOC,@StorageBin,@Operator,@IsTraceable,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
        Else
            NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RTLot,CreatedBy,CreatedOn,RecDate,Printed,ProdDate,RoHS,ReasonCode, StockType,MaterialDesc,SLOC,StorageBin,Operator,IsTraceable,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RTLot,@CreatedBy,getdate(),getdate(),@Printed,@ProdDate,@RoHS,@ReasonCode,@StockType,@MaterialDesc,@SLOC,@StorageBin,@Operator,@IsTraceable,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)
        End If
        'NewCLIDCommand = New SqlClient.SqlCommand("Insert into T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RTLot,CreatedBy,CreatedOn,RecDate,Printed,ProdDate,RoHS,ReasonCode, StockType,MaterialDesc,SLOC,StorageBin,Operator,IsTraceable,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,LastTransaction ) values (@NewCLID, @OrgCode, 1, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RTLot,@CreatedBy,getdate(),getdate(),@Printed,@ProdDate,@RoHS,@ReasonCode,@StockType,@MaterialDesc,@SLOC,@StorageBin,@Operator,@IsTraceable,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@LastTransaction ) ", myConn)

        NewCLIDCommand.CommandTimeout = TimeOut_M5

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
        NewCLIDCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
        NewCLIDCommand.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
        NewCLIDCommand.Parameters.Add("@Printed", SqlDbType.VarChar, 100, "Printed")

        NewCLIDCommand.Parameters.Add("@ProdDate", SqlDbType.SmallDateTime, 4, "ProdDate")
        NewCLIDCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
        NewCLIDCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 10, "StockType")
        NewCLIDCommand.Parameters.Add("@MaterialDesc", SqlDbType.VarChar, 250, "MaterialDesc")
        NewCLIDCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 20, "SLOC")
        NewCLIDCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
        NewCLIDCommand.Parameters.Add("@Operator", SqlDbType.VarChar, 10, "Operator")
        NewCLIDCommand.Parameters.Add("@IsTraceable", SqlDbType.VarChar, 10, "IsTraceable")
        NewCLIDCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 50, "Manufacturer")
        NewCLIDCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 50, "ManufacturerPN")
        NewCLIDCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
        NewCLIDCommand.Parameters.Add("@RoHS", SqlDbType.VarChar, 10, "RoHS")
        NewCLIDCommand.Parameters.Add("@AddlData", SqlDbType.VarChar, 20, "AddlData")
        NewCLIDCommand.Parameters.Add("@Stemp", SqlDbType.VarChar, 50, "Stemp")
        NewCLIDCommand.Parameters.Add("@MSL", SqlDbType.VarChar, 50, "MSL")
        NewCLIDCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 100, "LastTransaction")

        NewCLIDCommand.Parameters("@OrgCode").Value = LoginData.OrgCode
        NewCLIDCommand.Parameters("@MaterialNo").Value = LabelInfo.Material
        NewCLIDCommand.Parameters("@MaterialDesc").Value = LabelInfo.ItemData.Tables(0).Rows(0)("item_desc")
        NewCLIDCommand.Parameters("@MaterialRevision").Value = LabelInfo.MatlRev.ToString
        NewCLIDCommand.Parameters("@Qty").Value = LabelInfo.Qty
        NewCLIDCommand.Parameters("@UOM").Value = LabelInfo.ItemData.Tables(0).Rows(0)("uom_code")
        NewCLIDCommand.Parameters("@QtyBaseUOM").Value = LabelInfo.Qty
        NewCLIDCommand.Parameters("@BaseUOM").Value = LabelInfo.ItemData.Tables(0).Rows(0)("uom_code")
        NewCLIDCommand.Parameters("@DateCode").Value = LabelInfo.DCode
        NewCLIDCommand.Parameters("@LotNo").Value = LabelInfo.LotNo
        NewCLIDCommand.Parameters("@COO").Value = DBNull.Value
        NewCLIDCommand.Parameters("@RecDocNo").Value = LabelInfo.RTLot
        NewCLIDCommand.Parameters("@RTLot").Value = LabelInfo.RTLot
        NewCLIDCommand.Parameters("@CreatedBy").Value = LoginData.User.ToUpper
        NewCLIDCommand.Parameters("@SLOC").Value = LabelInfo.SubInv
        NewCLIDCommand.Parameters("@StorageBin").Value = LabelInfo.Locator
        NewCLIDCommand.Parameters("@Operator").Value = LoginData.User.ToUpper
        NewCLIDCommand.Parameters("@Printed").Value = lblPrint
        NewCLIDCommand.Parameters("@IsTraceable").Value = "NT"
        NewCLIDCommand.Parameters("@ReasonCode").Value = "CLID Generation"
        ' NewCLIDCommand.Parameters("@ExpDate").Value = ""   'DBNull.Value
        NewCLIDCommand.Parameters("@ProdDate").Value = DBNull.Value
        NewCLIDCommand.Parameters("@Manufacturer").Value = DBNull.Value
        NewCLIDCommand.Parameters("@ManufacturerPN").Value = DBNull.Value
        NewCLIDCommand.Parameters("@QMLStatus").Value = DBNull.Value
        NewCLIDCommand.Parameters("@RoHS").Value = DBNull.Value
        NewCLIDCommand.Parameters("@AddlData").Value = DBNull.Value
        NewCLIDCommand.Parameters("@Stemp").Value = DBNull.Value
        NewCLIDCommand.Parameters("@MSL").Value = DBNull.Value
        NewCLIDCommand.Parameters("@LastTransaction").Value = "CLID Generation"

        If LabelInfo.ItemData.Tables(0).Rows(0)("routing_id").ToString = "2" Then
            NewCLIDCommand.Parameters("@StockType").Value = "Q"
        Else
            NewCLIDCommand.Parameters("@StockType").Value = "FTS"
        End If

        'If LabelInfo.ExpDate Is Nothing OrElse LabelInfo.ExpDate Is DBNull.Value Then
        If LabelInfo.ExpDate <> "N" Then
            ' Get valid Expiration Date                  Format: MM/dd/yyyy
            Dim ExpDate As String
            Dim ExpiryDate As DateTime
            Dim ProdDate As DateTime
            NewCLIDCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
            ExpDate = Replace(LabelInfo.ExpDate, "-", "/")
            'ExpiryDate = CDate(ExpDate)

            NewCLIDCommand.Parameters("@ExpDate").Value = CDate(ExpDate)

            'Dim ExpDays As Double = 0
            'If Not LabelInfo.ItemData.Tables(0).Rows(0)("shelf_life_days") Is DBNull.Value Then
            '    ExpDays = LabelInfo.ItemData.Tables(0).Rows(0)("shelf_life_days")
            'End If
            'If ExpDays <> 0 Then
            '    ProdDate = ExpiryDate.AddDays(ExpDays * (-1))
            '    NewCLIDCommand.Parameters("@ProdDate").Value = ProdDate
            'End If
        End If
        If AMLFlag = True Then
            Dim Manufacturer, ManufacturerPN As String
            Manufacturer = LabelInfo.MFR
            ManufacturerPN = LabelInfo.MPN
            NewCLIDCommand.Parameters("@Manufacturer").Value = LabelInfo.MFR
            NewCLIDCommand.Parameters("@ManufacturerPN").Value = LabelInfo.MPN

            Dim DR() As DataRow = Nothing
            DR = dsAML.Tables("AMLData").Select(" MFR = '" & Manufacturer & "' and MPN = '" & ManufacturerPN & "'")
            If DR.Length > 0 Then
                NewCLIDCommand.Parameters("@QMLStatus").Value = DR(0)("AMLStatus").ToString
            Else
                NewCLIDCommand.Parameters("@QMLStatus").Value = "USERINPUT"
            End If
        End If

        If RoHSFlag = True Then
            NewCLIDCommand.Parameters("@RoHS").Value = dsAML.Tables("ItemData").Rows(0)("RoHS").ToString
            NewCLIDCommand.Parameters("@AddlData").Value = dsAML.Tables("ItemData").Rows(0)("AddlData").ToString
            NewCLIDCommand.Parameters("@Stemp").Value = dsAML.Tables("ItemData").Rows(0)("Stemp").ToString
            NewCLIDCommand.Parameters("@MSL").Value = dsAML.Tables("ItemData").Rows(0)("MSL").ToString
        End If

        Dim i, k As Integer

        Try
            myConn.Open()

            For i = 1 To CInt(LabelInfo.Units)
                'Decide CLID type according to item type
                myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
                myCommand.CommandType = CommandType.StoredProcedure
                myCommand.CommandTimeout = TimeOut_M5
                myCommand.Parameters.AddWithValue("@NextNo", "")
                myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
                myCommand.Parameters(0).Size = 20
                myCommand.Parameters(0).Direction = ParameterDirection.Output
                If LabelInfo.ItemData.Tables(0).Rows(0)("type_name") = "RM" Then
                    myCommand.Parameters.AddWithValue("@TypeID", "CLID")
                Else
                    myCommand.Parameters.AddWithValue("@TypeID", "PCBID")
                End If

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
                    NewCLIDCommand.Parameters("@NewCLID").Value = NextCLID
                    NewCLIDCommand.CommandType = CommandType.Text
                    ra = NewCLIDCommand.ExecuteNonQuery()

                    myDataRow = CLIDsTable.NewRow()
                    myDataRow("CLID") = NextCLID
                    CLIDsTable.Rows.Add(myDataRow)

                End If

            Next

            LabelConversion.CLIDs = New DataSet
            LabelConversion.CLIDs = CLIDs

            If LabelConversion.CLIDs Is Nothing OrElse LabelConversion.CLIDs.Tables.Count = 0 OrElse LabelConversion.CLIDs.Tables(0).Rows.Count = 0 Then
                LabelConversion.ErrorMsg = "CLID Generation error"
            End If

            myConn.Close()

        Catch ex As Exception
            ErrorLogging("LabelBasicFunction-LabelConversion", LoginData.User.ToUpper, "RTNo: " & LabelInfo.RTNo & ", " & ex.Message & ex.Source, "E")
            LabelConversion.ErrorMsg = "CLID Generation error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return LabelConversion
    End Function

    Public Function GetRTVLabels(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal CVMIFlag As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            GetRTVLabels = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_GetRTVLabels  '{0}','{1}' ,'{2}' ", LoginData.OrgCode, CLID, CVMIFlag)
                GetRTVLabels = da.ExecuteDataSet(Sqlstr, "LabelIDs")
                Return GetRTVLabels

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-GetRTVLabels", LoginData.User.ToUpper, "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function MiscRTV(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i, j As Integer
                Dim DR() As DataRow = Nothing
                Dim ds As DataSet = New DataSet
                Dim Sqlstr, MaterialNo, MergedID, RTLot As String

                'Get Original CLIDs and TransactionID for the MergedIDs if there exists, need to handle MergedID to do CVMI RTV         
                DR = Items.Tables(1).Select(" TransactionID = ' ' or TransactionID IS NULL ")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        RTLot = DR(i)("RTLot").ToString
                        MaterialNo = DR(i)("MaterialNo").ToString
                        MergedID = DR(i)("CLID").ToString

                        Dim SubInv, Locator As String
                        SubInv = DR(i)("SLOC").ToString
                        Locator = DR(i)("StorageBin").ToString

                        ds = New DataSet
                        Sqlstr = String.Format("Select CLID,OrgCode,MaterialNo,MaterialRevision,QtyBaseUOM,BaseUOM,StatusCode,VendorID,PurOrdNo,PurOrdItem,SLOC, StorageBin, RTLot, CONVERT(varchar, ExpDate, 101) as Expdate,RecDocNo as RTNo, RecDocItem as TransactionID, BoxID,ReasonCode='' from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and StatusCode = 0 and OrgCode ='{0}' and MaterialNo = '{1}' and RTLot = '{2}' and ReferenceCLID ='{3}' ", LoginData.OrgCode, MaterialNo, RTLot, MergedID)
                        ds = da.ExecuteDataSet(Sqlstr, "RTVData")
                        If ds.Tables(0).Rows.Count = 0 Then
                            MiscRTV = "CLID " & MergedID & " was created by spliting through MergeID, could not find the Original CLIDs, not allow to do RTV"
                            Exit Function
                        End If

                        'Get the correct Source SubInv / Locator for Original CLIDs
                        For j = 0 To ds.Tables(0).Rows.Count - 1
                            ds.Tables(0).Rows(j)("SLOC") = SubInv
                            ds.Tables(0).Rows(j)("StorageBin") = Locator
                        Next
                        ds.Tables(0).AcceptChanges()
                        Items.Tables(1).Merge(ds.Tables(0))
                    Next
                End If

                Dim RTVData As DataTable = Items.Tables(1).Clone()
                RTVData.Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
                ds = New DataSet
                ds.Tables.Add(RTVData)

                Dim DestSubInv As String = Items.Tables(0).Rows(0)("DestSubInv").ToString
                Dim DestLocator As String = Items.Tables(0).Rows(0)("DestLocator").ToString
                Dim RTVNumber As String = Items.Tables(0).Rows(0)("RTVNumber").ToString
                Dim RMANumber As String = Items.Tables(0).Rows(0)("RMANumber").ToString

                j = -1
                For i = 0 To Items.Tables(1).Rows.Count - 1
                    Dim TransactionID As String = Items.Tables(1).Rows(i)("TransactionID").ToString
                    If TransactionID = "" Then Continue For

                    Dim SubInv, Locator, PurOrdNo, PurOrdItem As String
                    SubInv = Items.Tables(1).Rows(i)("SLOC").ToString
                    Locator = Items.Tables(1).Rows(i)("StorageBin").ToString
                    RTLot = Items.Tables(1).Rows(i)("RTLot").ToString
                    PurOrdNo = Items.Tables(1).Rows(i)("PurOrdNo")
                    PurOrdItem = Items.Tables(1).Rows(i)("PurOrdItem")

                    If Items.Tables(1).Rows(i)("ExpDate") Is DBNull.Value Then
                        Sqlstr = " SLOC = '" & SubInv & "' and StorageBin = '" & Locator & "' and RTLot = '" & RTLot & "' and PurOrdNo = '" & PurOrdNo & "' and PurOrdItem = '" & PurOrdItem & "'"
                    Else
                        Sqlstr = " SLOC = '" & SubInv & "' and StorageBin = '" & Locator & "' and RTLot = '" & RTLot & "' and ExpDate = '" & Items.Tables(1).Rows(i)("ExpDate") & "' and PurOrdNo = '" & PurOrdNo & "' and PurOrdItem = '" & PurOrdItem & "'"
                    End If

                    DR = Nothing
                    DR = ds.Tables(0).Select(Sqlstr)
                    If DR.Length = 0 Then
                        j = j + 1
                        ds.Tables(0).ImportRow(Items.Tables(1).Rows(i))
                        ds.Tables(0).Rows(j)("CLID") = ""
                        ds.Tables(0).Rows(j)("BoxID") = ""
                        ds.Tables(0).Rows(j)("DestSubInv") = DestSubInv
                        ds.Tables(0).Rows(j)("DestLocator") = DestLocator
                        ds.Tables(0).Rows(j)("RTVNumber") = RTVNumber
                        ds.Tables(0).Rows(j)("RMANumber") = RMANumber
                        ds.Tables(0).Rows(j)("Status") = ""
                        ds.Tables(0).Rows(j)("Message") = ""

                        ds.Tables(0).Rows(j).AcceptChanges()
                        ds.Tables(0).Rows(j).SetAdded()
                    Else
                        Dim Qty As Decimal = DR(0)("QtyBaseUOM")
                        DR(0)("QtyBaseUOM") = Qty + Items.Tables(1).Rows(i)("QtyBaseUOM")
                        DR(0).AcceptChanges()
                        DR(0).SetAdded()
                    End If
                Next


                Try
                    MiscRTV = Oracle_RTVTransfer(LoginData, ds)
                    If MiscRTV <> "Y" Then Exit Function
                Catch ex As Exception
                    ErrorLogging("LabelBasicFunction-MiscRTV1", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                    Return ex.Message
                End Try


                Try
                    Dim CVMIFlag As String = "N"
                    Dim dsItem As DataSet = New DataSet

                    dsItem.DataSetName = "Items"
                    dsItem.Merge(Items.Tables(0))

                    Dim ProcessRTV As String = ""
                    Sqlstr = String.Format("exec sp_ProcessRTV '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, LoginData.User.ToUpper, CVMIFlag, DStoXML(dsItem))
                    ProcessRTV = da.ExecuteScalar(Sqlstr).ToString
                    Return ProcessRTV
                Catch ex As Exception
                    ErrorLogging("MiscRTV-sp_ProcessRTV", LoginData.User, ex.Message & ex.Source, "E")
                    MiscRTV = "Data update error, please contact IT."
                End Try

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-MiscRTV", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                MiscRTV = ex.Message
            End Try

        End Using

    End Function

    Public Function Oracle_RTVTransfer(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()
            Dim TransactionID As Integer
            Dim s_date As String = Format(DateTime.Now, "MMddHHmmss")
            TransactionID = CInt(s_date) + CInt(LoginData.UserID)

            Dim aa As OracleString
            Dim oda As OracleDataAdapter = New OracleDataAdapter()
            Dim comm As OracleCommand = da.Ora_Command_Trans()

            Try
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
                comm.Parameters.Add("p_organization_code", OracleType.VarChar, 240).Value = OrgID  'LoginData.OrgCode
                comm.Parameters.Add("p_transaction_header_id", OracleType.Double).Value = TransactionID
                comm.Parameters.Add("p_transaction_type_name", OracleType.VarChar, 240).Value = "Return to Vendor Transfer"
                'comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240).Value = "Return to Vendor"
                comm.Parameters.Add("p_reason_code", OracleType.VarChar, 240).Value = "ST-RTV"
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
                'comm.Parameters.Add("p_transaction_reference", OracleType.VarChar, 240)
                comm.Parameters.Add("p_vendor_number", OracleType.VarChar, 240)
                comm.Parameters.Add("p_po_number", OracleType.VarChar, 240)
                comm.Parameters.Add("p_po_line", OracleType.VarChar, 240)
                comm.Parameters.Add("p_rtv_number", OracleType.VarChar, 240)
                comm.Parameters.Add("p_rma_number", OracleType.VarChar, 240)

                'Use srf_number to save Receipt Number in MTL_MATERIAL_TRANSACTIONS.ATTRIBUTE13, which requested by RexWong 06/14/2012
                comm.Parameters.Add("p_srf_number", OracleType.VarChar, 240)

                ''/////////////////////////////////////////////////////////////////////////////////
                ''//Comments by Jackson
                ''//Date:03/31/2011
                ''//
                ''//Due to the user no need add these two fieids in RTV program
                ''//But Oracle already have these two fields so if we want to use these two fields
                ''//We can cancel comment will be ok
                ''//
                ''/////////////////////////////////////////////////////////////////////////////////

                ''comm.Parameters.Add("p_srf_number", OracleType.VarChar, 240)
                ''comm.Parameters.Add("p_chargeable", OracleType.VarChar, 240)

                '                    Select CLID,OrgCode,MaterialNo,MaterialRevision,QtyBaseUOM,BaseUOM,StatusCode,VendorID,PurOrdNo,PurOrdItem,SLOC, 
                'StorageBin, RTLot, CONVERT(varchar, ExpDate, 101) as Expdate,RecDocNo as RTNo, RecDocItem as TransactionID, BoxID,ReasonCode='' from T_CLMaster with (nolock) where StatusCode = 1 and (SLOC like '%MRB%' or SLOC like '%MR1') and PurOrdItem <> '' and PurOrdNo <> '' and OrgCode = @OrgCode and MaterialNo = @MaterialNo and VendorID = @VendorID
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
                'comm.Parameters("p_transaction_reference").SourceColumn = "p_transaction_reference"
                comm.Parameters("p_vendor_number").SourceColumn = "VendorID"
                comm.Parameters("p_po_number").SourceColumn = "PurOrdNo"
                comm.Parameters("p_po_line").SourceColumn = "PurOrdItem"
                comm.Parameters("p_rtv_number").SourceColumn = "RTVNumber"
                comm.Parameters("p_rma_number").SourceColumn = "RMANumber"

                'Use srf_number to save Receipt Number in MTL_MATERIAL_TRANSACTIONS.ATTRIBUTE13, which requested by RexWong 06/14/2012
                comm.Parameters("p_srf_number").SourceColumn = "RTNo"

                ''/////////////////////////////////////////////////////////////////////////////////
                ''//Comments by Jackson
                ''//Date:03/31/2011
                ''//
                ''//Due to the user no need add these two fieids in RTV program
                ''//But Oracle already have these two fields so if we want to use these two fields
                ''//We can cancel comment will be ok
                ''//
                ''/////////////////////////////////////////////////////////////////////////////////

                ''comm.Parameters("p_srf_number").SourceColumn = "SRFNumber"
                ''comm.Parameters("p_chargeable").SourceColumn = "Chargeable"

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
                    Exit Function
                End If

                'Rollback if there has error
                comm.Transaction.Rollback()
                comm.Connection.Close()

                'Record error message to eTrace database
                Dim i As Integer

                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString.Trim <> "" Then
                        If Msg = "" Then Msg = DR(i)("Message").ToString.Trim
                        Dim ErrMsg As String = "RTLot: " & DR(i)("RTLot").ToString & " with Item " & DR(i)("MaterialNo") & " and return message: "
                        ErrMsg = ErrMsg & DR(i)("Message").ToString.Trim
                        ErrorLogging("LabelBasicFunction-Oracle_RTVTransfer1", LoginData.User.ToUpper, ErrMsg, "I")
                    End If
                Next
                Return Msg

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-Oracle_RTVTransfer", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                comm.Transaction.Rollback()
                Return "RTV Transaction error"
            Finally
                If comm.Connection.State <> ConnectionState.Closed Then comm.Connection.Close()
            End Try
        End Using

    End Function

    Public Function CVMIRTV(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i, j As Integer
                Dim DR() As DataRow = Nothing
                Dim ds As DataSet = New DataSet
                Dim Sqlstr, MaterialNo, MergedID, RTLot As String

                'Get Original CLIDs and TransactionID for the MergedIDs if there exists, need to handle MergedID to do CVMI RTV         
                DR = Items.Tables(1).Select(" TransactionID = ' ' or TransactionID IS NULL ")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        RTLot = DR(i)("RTLot").ToString
                        MaterialNo = DR(i)("MaterialNo").ToString
                        MergedID = DR(i)("CLID").ToString

                        ds = New DataSet
                        Sqlstr = String.Format("Select CLID, OrgCode, MaterialNo,QtyBaseUOM,BaseUOM,StatusCode,VendorID,PurOrdNo,PurOrdItem,SLOC,StorageBin, RTLot,RecDocNo as RTNo, RecDocItem as TransactionID, BoxID, ReasonCode='', TotalQty=0.000 from T_CLMaster with (nolock) where (SLOC IS NULL or SLOC = '') and StatusCode = 0 and OrgCode ='{0}' and MaterialNo = '{1}' and RTLot = '{2}' and ReferenceCLID ='{3}' ", LoginData.OrgCode, MaterialNo, RTLot, MergedID)
                        ds = da.ExecuteDataSet(Sqlstr, "RTVData")
                        If ds.Tables(0).Rows.Count = 0 Then
                            CVMIRTV = "CLID " & MergedID & " was created by spliting through MergeID, could not find the Original CLIDs, not allow to do CVMI RTV"
                            Exit Function
                        End If

                        Items.Tables(1).Merge(ds.Tables(0))
                    Next
                End If

                Dim RTVData As DataTable = Items.Tables(1).Clone()
                RTVData.Columns.Add(New Data.DataColumn("PONo", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("ReleaseNo", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("POLine", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("ShipmentNo", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("ValidFlag", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
                RTVData.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
                ds = New DataSet
                ds.Tables.Add(RTVData)

                Dim SubInv As String = Items.Tables(0).Rows(0)("SLOC")
                Dim Locator As String = Items.Tables(0).Rows(0)("StorageBin")
                Dim TotalQty As Decimal = Items.Tables(0).Rows(0)("TotalQty")
                Dim ReasonCode As String = Items.Tables(0).Rows(0)("ReasonCode")

                j = -1
                For i = 0 To Items.Tables(1).Rows.Count - 1
                    Dim TransactionID As String = Items.Tables(1).Rows(i)("TransactionID").ToString
                    If TransactionID = "" Then Continue For

                    Dim PurOrdNo, PurOrdItem As String
                    PurOrdNo = Items.Tables(1).Rows(i)("PurOrdNo")
                    PurOrdItem = Items.Tables(1).Rows(i)("PurOrdItem")

                    DR = Nothing
                    'DR = ds.Tables(0).Select(" PurOrdNo = '" & PurOrdNo & "' and PurOrdItem = '" & PurOrdItem & "'")
                    DR = ds.Tables(0).Select(" TransactionID = '" & TransactionID & "'")
                    If DR.Length = 0 Then
                        Dim OrderData As POData
                        Dim myReceiving As Receiving = New Receiving
                        OrderData = myReceiving.Split_POData(PurOrdNo, PurOrdItem)

                        j = j + 1
                        ds.Tables(0).ImportRow(Items.Tables(1).Rows(i))
                        ds.Tables(0).Rows(j)("CLID") = ""
                        ds.Tables(0).Rows(j)("BoxID") = ""
                        ds.Tables(0).Rows(j)("PONo") = OrderData.PONo
                        ds.Tables(0).Rows(j)("POLine") = OrderData.LineNo
                        ds.Tables(0).Rows(j)("ShipmentNo") = OrderData.ShipmentNo
                        ds.Tables(0).Rows(j)("ReleaseNo") = OrderData.ReleaseNo
                        If OrderData.ReleaseNo = 0 Then ds.Tables(0).Rows(j)("ReleaseNo") = ""
                        If FixNull(ds.Tables(0).Rows(j)("SLOC")) = "" Then ds.Tables(0).Rows(j)("SLOC") = SubInv
                        If FixNull(ds.Tables(0).Rows(j)("StorageBin")) = "" Then ds.Tables(0).Rows(j)("StorageBin") = Locator

                        ds.Tables(0).Rows(j)("ReasonCode") = ReasonCode
                        ds.Tables(0).Rows(j)("TotalQty") = TotalQty
                        ds.Tables(0).Rows(j)("Status") = ""
                        ds.Tables(0).Rows(j)("Message") = ""
                        ds.Tables(0).Rows(j)("ValidFlag") = "N"
                        If j = 0 Then ds.Tables(0).Rows(j)("ValidFlag") = "Y"

                        ds.Tables(0).Rows(j).AcceptChanges()
                        ds.Tables(0).Rows(j).SetAdded()
                    Else
                        Dim Qty As Decimal = DR(0)("QtyBaseUOM")
                        DR(0)("QtyBaseUOM") = Qty + Items.Tables(1).Rows(i)("QtyBaseUOM")
                        DR(0).AcceptChanges()
                        DR(0).SetAdded()
                    End If
                Next

                Try
                    Dim UserID As String = LoginData.User.ToUpper
                    CVMIRTV = Oracle_CVMIRTV(LoginData, ds)
                    If CVMIRTV <> "Y" Then Exit Function
                Catch ex As Exception
                    ErrorLogging("LabelBasicFunction-CVMIRTV1", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                    Return ex.Message
                End Try

                Try
                    Dim CVMIFlag As String = "Y"
                    Dim dsItem As DataSet = New DataSet

                    dsItem.DataSetName = "Items"
                    dsItem.Merge(Items.Tables(0))
                    dsItem.Tables(0).Columns.Add(New Data.DataColumn("DestSubInv", System.Type.GetType("System.String")))
                    dsItem.Tables(0).Columns.Add(New Data.DataColumn("DestLocator", System.Type.GetType("System.String")))

                    Dim ProcessRTV As String = ""
                    Sqlstr = String.Format("exec sp_ProcessRTV '{0}','{1}','{2}',N'{3}'", LoginData.OrgCode, LoginData.User.ToUpper, CVMIFlag, DStoXML(dsItem))
                    ProcessRTV = da.ExecuteScalar(Sqlstr).ToString
                    Return ProcessRTV
                Catch ex As Exception
                    ErrorLogging("CVMIRTV-sp_ProcessRTV", LoginData.User, ex.Message & ex.Source, "E")
                    CVMIRTV = "Data update error, please contact IT."
                End Try

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-CVMIRTV", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                CVMIRTV = ex.Message
            End Try
        End Using

    End Function

    Public Function Oracle_CVMIRTV(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As OracleString
            Dim oc As OracleCommand = da.Ora_Command_Trans()
            Dim oda As OracleDataAdapter = New OracleDataAdapter()

            Dim RTNo As String = p_ds.Tables(0).Rows(0)("RTNo").ToString

            Try

                Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables(0).Rows(0)("PONo"), 4, 2)
                If POStr = "04" Then LoginData.RespID_Inv = GetGlobalRespID(LoginData.Server)

                oc.CommandType = CommandType.StoredProcedure
                oc.CommandText = "apps.xxetr_wip_pkg.initialize"
                oc.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                oc.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
                oc.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID
                oc.ExecuteOracleNonQuery(aa)
                oc.Parameters.Clear()

                oc.CommandText = "apps.xxetr_rcv_return_pkg.process_vmi_return"       'CVMI PO Reversal
                oc.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = LoginData.OrgCode
                oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
                oc.Parameters.Add("p_po_number", OracleType.VarChar, 50)
                oc.Parameters.Add("p_receipt_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_line_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_shipment_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_release_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_quantity", OracleType.Double)
                oc.Parameters.Add("p_total_qty", OracleType.Double)
                oc.Parameters.Add("p_uom_code", OracleType.VarChar, 50)
                oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
                oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
                oc.Parameters.Add("p_lot_num", OracleType.VarChar, 50)
                oc.Parameters.Add("p_reason_code", OracleType.VarChar, 200)
                oc.Parameters.Add("p_valid_flag", OracleType.VarChar, 50)
                oc.Parameters.Add("p_rma_number", OracleType.VarChar, 500).Value = "eTraceHH CVMI RTV"
                oc.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                oc.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                oc.Parameters("p_transaction_id").SourceColumn = "TransactionID"                      'TransactionID           
                oc.Parameters("p_po_number").SourceColumn = "PONo"
                oc.Parameters("p_receipt_num").SourceColumn = "RTNo"
                oc.Parameters("p_line_num").SourceColumn = "POLine"
                oc.Parameters("p_shipment_num").SourceColumn = "ShipmentNo"
                oc.Parameters("p_release_num").SourceColumn = "ReleaseNo"
                oc.Parameters("p_item_num").SourceColumn = "MaterialNo"
                oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
                oc.Parameters("p_total_qty").SourceColumn = "TotalQty"
                oc.Parameters("p_uom_code").SourceColumn = "BaseUOM"
                oc.Parameters("p_subinv").SourceColumn = "SLOC"
                oc.Parameters("p_locator").SourceColumn = "StorageBin"
                oc.Parameters("p_lot_num").SourceColumn = "RTLot"
                oc.Parameters("p_reason_code").SourceColumn = "ReasonCode"
                oc.Parameters("p_valid_flag").SourceColumn = "ValidFlag"
                oc.Parameters("o_success_flag").SourceColumn = "Status"
                oc.Parameters("o_error_mssg").SourceColumn = "Message"

                oda.InsertCommand = oc
                oda.Update(p_ds.Tables(0))

                Dim Flag, Msg As String
                Flag = oc.Parameters("o_success_flag").Value.ToString
                Msg = oc.Parameters("o_error_mssg").Value.ToString

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables(0).Select("Status = 'N' or Status = ' '")
                If DR.Length = 0 Then
                    oc.Transaction.Commit()
                    oc.Connection.Close()
                    Return DirectCast(Flag, String)
                    Exit Function
                End If

                'Rollback if there has error
                oc.Transaction.Rollback()
                oc.Connection.Close()

                'Record error message to eTrace database
                Dim i As Integer
                Dim ErrMsg As String = "RTNo: " & RTNo & " with Item " & DR(0)("MaterialNo") & " and return message "
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString.Trim <> "" Then
                        If Msg = "" Then Msg = DR(i)("Message").ToString.Trim
                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString.Trim
                    End If
                Next
                ErrorLogging("LabelBasicFunction-Oracle_CVMIRTV1", LoginData.User.ToUpper, ErrMsg, "I")
                Return Msg

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Oracle_CVMIRTV", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                oc.Transaction.Rollback()
                Return "CVMI RTV Transaction error"
            Finally
                If oc.Connection.State <> ConnectionState.Closed Then oc.Connection.Close()
            End Try

        End Using
    End Function

    Public Function GetItemRTLot(ByVal LoginData As ERPLogin, ByVal PartNo As String, ByVal RTLot As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try

                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_po_detail_pkg.validate_rt_item"
                OC.Parameters.Add("p_org_code", OracleType.VarChar).Value = LoginData.OrgCode
                OC.Parameters.Add("p_item_num", OracleType.VarChar, 20).Value = PartNo.ToUpper()
                OC.Parameters.Add("p_rt_lot", OracleType.VarChar, 50).Value = RTLot.ToUpper()
                OC.Parameters.Add("o_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_exp_date", OracleType.VarChar, 100).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Dim Flag, ExpDate As String
                Flag = OC.Parameters("o_flag").Value.ToString
                ExpDate = OC.Parameters("o_exp_date").Value.ToString

                OC.Connection.Close()
                'OC.Dispose()

                If Flag = "Y" Then        ' Y: Valid RTLot;  N: Invalid RTLot
                    GetItemRTLot = ExpDate
                Else
                    GetItemRTLot = Flag
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-GetItemRTLot", LoginData.User.ToUpper, "PartNo: " & PartNo & " with RTLot: " & RTLot & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetItemOnhand(ByVal LoginData As ERPLogin, ByVal ItemNo As String, ByVal Subinv As String, ByVal LocName As String, ByVal Revision As String, ByVal LotNum As String, ByVal Qty As String) As String()
        Using da As DataAccess = GetDataAccess()
            Dim l_success_flag(1) As String
            Dim oCommand As OracleCommand = da.OraCommand()
            Try
                If oCommand.Connection.State = ConnectionState.Closed Then
                    oCommand.Connection.Open()
                End If

                Dim rowCount As Integer
                oCommand.CommandType = CommandType.StoredProcedure
                oCommand.CommandText = "apps.xxetr_wip_pkg.p_get_item_qoh"
                If ItemNo Is DBNull.Value Then
                    oCommand.Parameters.Add("p_item_number", OracleType.VarChar, 60).Value = ""
                Else
                    oCommand.Parameters.Add("p_item_number", OracleType.VarChar, 60).Value = ItemNo
                End If
                If LoginData.OrgCode Is DBNull.Value Then
                    oCommand.Parameters.Add("p_org_code", OracleType.VarChar, 30).Value = ""
                Else
                    oCommand.Parameters.Add("p_org_code", OracleType.VarChar, 30).Value = LoginData.OrgCode
                End If
                If Subinv Is DBNull.Value Then
                    oCommand.Parameters.Add("p_subinventory_code", OracleType.VarChar, 30).Value = ""
                Else
                    oCommand.Parameters.Add("p_subinventory_code", OracleType.VarChar, 30).Value = Subinv
                End If
                If LocName Is DBNull.Value Then
                    oCommand.Parameters.Add("p_locator_name", OracleType.VarChar, 60).Value = ""
                Else
                    oCommand.Parameters.Add("p_locator_name", OracleType.VarChar, 60).Value = LocName
                End If
                If Revision Is DBNull.Value Then
                    oCommand.Parameters.Add("p_item_rev", OracleType.VarChar, 20).Value = ""
                Else
                    oCommand.Parameters.Add("p_item_rev", OracleType.VarChar, 20).Value = Revision
                End If
                If LotNum Is DBNull.Value Then
                    oCommand.Parameters.Add("p_lot_number", OracleType.VarChar, 20).Value = ""
                Else
                    oCommand.Parameters.Add("p_lot_number", OracleType.VarChar, 20).Value = LotNum
                End If
                oCommand.Parameters.Add("p_qty_onhand", OracleType.Number)
                oCommand.Parameters.Add("o_exp_date", OracleType.VarChar, 20)
                oCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 2)
                oCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 100)
                oCommand.Parameters("p_qty_onhand").Direction = ParameterDirection.Output
                oCommand.Parameters("o_exp_date").Direction = ParameterDirection.Output
                oCommand.Parameters("o_success_flag").Direction = ParameterDirection.Output
                oCommand.Parameters("o_error_mssg").Direction = ParameterDirection.Output

                rowCount = oCommand.ExecuteNonQuery()
                l_success_flag(0) = oCommand.Parameters("o_success_flag").Value.ToString()

                If l_success_flag(0) = "Y" Then
                    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
                    Dim cmdSQL As SqlClient.SqlCommand
                    myConn.Open()
                    cmdSQL = New SqlClient.SqlCommand("sp_GetOnHand", myConn)
                    cmdSQL.CommandType = CommandType.StoredProcedure
                    cmdSQL.CommandTimeout = TimeOut_M5

                    cmdSQL.Parameters.Add("@p_org", SqlDbType.VarChar, 15)
                    cmdSQL.Parameters("@p_org").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@p_org").Value = LoginData.OrgCode

                    cmdSQL.Parameters.Add("@p_item_num", SqlDbType.VarChar, 50)
                    cmdSQL.Parameters("@p_item_num").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@p_item_num").Value = ItemNo

                    cmdSQL.Parameters.Add("@p_version", SqlDbType.VarChar, 15)
                    cmdSQL.Parameters("@p_version").Direction = ParameterDirection.Input
                    If Revision Is DBNull.Value Then
                        cmdSQL.Parameters("@p_version").Value = ""
                    Else
                        cmdSQL.Parameters("@p_version").Value = Revision
                    End If

                    cmdSQL.Parameters.Add("@p_sub", SqlDbType.VarChar, 20)
                    cmdSQL.Parameters("@p_sub").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@p_sub").Value = Subinv

                    cmdSQL.Parameters.Add("@p_locator", SqlDbType.VarChar, 50)
                    cmdSQL.Parameters("@p_locator").Direction = ParameterDirection.Input
                    If LocName Is DBNull.Value Then
                        cmdSQL.Parameters("@p_locator").Value = ""
                    Else
                        cmdSQL.Parameters("@p_locator").Value = LocName
                    End If

                    cmdSQL.Parameters.Add("@p_rtlot", SqlDbType.VarChar, 20)
                    cmdSQL.Parameters("@p_rtlot").Direction = ParameterDirection.Input
                    If LotNum Is DBNull.Value Then
                        cmdSQL.Parameters("@p_rtlot").Value = ""
                    Else
                        cmdSQL.Parameters("@p_rtlot").Value = LotNum
                    End If

                    cmdSQL.Parameters.Add("@o_onhand", SqlDbType.Decimal, 20)
                    cmdSQL.Parameters("@o_onhand").Direction = ParameterDirection.Output
                    cmdSQL.ExecuteNonQuery()
                    If CDec(Qty) > (oCommand.Parameters("p_qty_onhand").Value - cmdSQL.Parameters("@o_onhand").Value) Then
                        l_success_flag(0) = "N"
                    End If
                    cmdSQL.Connection.Close()
                    'cmdSQL.Dispose()
                Else
                    l_success_flag(0) = oCommand.Parameters("o_error_mssg").Value.ToString
                End If
            Catch ex As Exception
                ErrorLogging("Receiving-GetItemOnhand", LoginData.User, ex.Message & ex.Source, "E")
                l_success_flag(0) = "N"
            Finally
                l_success_flag(1) = oCommand.Parameters("o_exp_date").Value
                If oCommand.Connection.State <> ConnectionState.Closed Then oCommand.Connection.Close()
                'oCommand.Dispose()
            End Try
            Return l_success_flag
        End Using
    End Function

    Public Function ReadLabelIDInfo(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal StatusCheck As Boolean) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim SQLString As String
        If StatusCheck = True Then
            SQLString = "SELECT CLID,MaterialNo, ISNULL(MaterialRevision,'') MaterialRevision, QtyBaseUOM, BaseUOM,ISNULL(RTLot,'') RTLot, ISNULL(SLOC,'') SLOC, ISNULL(StorageBin,'') StorageBin,ISNULL(Qty,0) Qty,ISNULL(UOM,'') UOM FROM T_CLMaster WHERE  OrgCode = '" & LoginData.OrgCode & "' AND CLID = '" & CLID & "' AND StatusCode = '1'"
        Else
            SQLString = "SELECT CLID,MaterialNo, ISNULL(MaterialRevision,'') MaterialRevision, QtyBaseUOM, BaseUOM,ISNULL(RTLot,'') RTLot, ISNULL(SLOC,'') SLOC, ISNULL(StorageBin,'') StorageBin,ISNULL(Qty,0) Qty,ISNULL(UOM,'') UOM FROM T_CLMaster WHERE  OrgCode = '" & LoginData.OrgCode & "' AND CLID = '" & CLID & "' AND StatusCode = '0'"
        End If

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(SQLString, myConn)
            sda.SelectCommand.CommandType = CommandType.Text
            sda.Fill(dsmodel, "model")
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetModelRevision", LoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function LabelGeneration(ByVal LoginData As ERPLogin, ByVal CLID As String, ByVal QtyUnits As String, ByVal Type As String) As String()
        Dim CLIDStr() As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_LabelGeneration", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@OrgCode").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@OrgCode").Value = LoginData.OrgCode

            cmdSQL.Parameters.Add("@CLID", SqlDbType.VarChar, 8000)
            cmdSQL.Parameters("@CLID").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@CLID").Value = CLID

            cmdSQL.Parameters.Add("@QtyUnits", SqlDbType.VarChar, 8000)
            cmdSQL.Parameters("@QtyUnits").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@QtyUnits").Value = QtyUnits

            cmdSQL.Parameters.Add("@User", SqlDbType.VarChar, 30)
            cmdSQL.Parameters("@User").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@User").Value = LoginData.User

            cmdSQL.Parameters.Add("@Type", SqlDbType.VarChar, 30)
            cmdSQL.Parameters("@Type").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Type").Value = Type

            cmdSQL.Parameters.Add("@newCLID", SqlDbType.VarChar, 8000)
            cmdSQL.Parameters("@newCLID").Direction = ParameterDirection.Output

            cmdSQL.ExecuteNonQuery()
            CLIDStr = Split(cmdSQL.Parameters("@newCLID").Value, ",")
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("LabelGeneration", LoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return CLIDStr
    End Function

    Public Function GetTypeID(ByVal CLID As String) As String
        Dim typeID, result As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_GetTypeID", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@p_clid", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@p_clid").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@p_clid").Value = CLID

            cmdSQL.Parameters.Add("@o_type", SqlDbType.VarChar, 25)
            cmdSQL.Parameters("@o_type").Direction = ParameterDirection.Output

            result = cmdSQL.ExecuteScalar()
            myConn.Close()
            If result = "1" Then
                typeID = cmdSQL.Parameters("@o_type").Value
            Else
                typeID = "Error"
            End If
        Catch ex As Exception
            ErrorLogging("GetTypeID", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return typeID
    End Function

    Public Function ReadLabelIDsForDJ(ByVal LoginData As ERPLogin, ByVal CLID As String) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim SQLString As String
        SQLString = "SELECT CLID,MaterialNo, ISNULL(MaterialRevision,'') MaterialRevision, QtyBaseUOM, BaseUOM,ISNULL(RTLot,'') RTLot, SLOC, ISNULL(StorageBin,'') StorageBin FROM T_CLMaster WHERE  OrgCode = '" & LoginData.OrgCode & "' AND CLID = '" & CLID & "' AND SLOC LIKE '%BF%'"

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        Try
            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(SQLString, myConn)
            sda.SelectCommand.CommandType = CommandType.Text
            sda.Fill(dsmodel, "LabelIDs")
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("ReadLabelIDsForDJ", LoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function Get_LabelInfoHasSameLot(ByVal CLID As String, ByVal LoginData As ERPLogin) As DataSet
        Dim cmdSQL As SqlClient.SqlCommand
        Dim adapter As SqlClient.SqlDataAdapter
        Dim ds As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_getHasSameLotCLID", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure

            cmdSQL.Parameters.Add("@Org", SqlDbType.VarChar, 20)
            cmdSQL.Parameters("@Org").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Org").Value = LoginData.OrgCode

            cmdSQL.Parameters.Add("@CLID", SqlDbType.VarChar, 60)
            cmdSQL.Parameters("@CLID").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@CLID").Value = CLID
            adapter = New SqlClient.SqlDataAdapter(cmdSQL)
            adapter.Fill(ds)
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Get_LabelInfoHasSameLot", LoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return ds
    End Function

    Public Function Update_LabelInfoHasSameLot(ByVal LoginData As ERPLogin, ByVal ItemNO As String, ByVal LotNum As String, ByVal ExpDate As Date) As String
        Using da As DataAccess = GetDataAccess()
            Dim oCommand As OracleCommand = da.OraCommand()
            Dim x_status As String
            Dim x_msg As String = ""
            Try
                If oCommand.Connection.State = ConnectionState.Closed Then
                    oCommand.Connection.Open()
                End If
                oCommand.CommandType = CommandType.StoredProcedure
                oCommand.CommandText = "apps.xxetr_update_mtl_attributes.update_mtl_expiration_date"

                oCommand.Parameters.Add("p_org_code", OracleType.VarChar, 20).Value = LoginData.OrgCode
                oCommand.Parameters.Add("p_item_segment1", OracleType.VarChar, 60).Value = ItemNO
                oCommand.Parameters.Add("p_lot_number", OracleType.VarChar, 60).Value = LotNum
                oCommand.Parameters.Add("p_exp_date", OracleType.DateTime).Value = ExpDate
                oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 40).Value = LoginData.User

                oCommand.Parameters.Add("x_status", OracleType.VarChar, 2)
                oCommand.Parameters.Add("x_msg", OracleType.VarChar, 500)

                oCommand.Parameters("x_status").Direction = ParameterDirection.Output
                oCommand.Parameters("x_msg").Direction = ParameterDirection.Output

                oCommand.ExecuteNonQuery()
                x_status = oCommand.Parameters("x_status").Value.ToString()
                If x_status = "1" Then
                    Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
                    Dim cmdSQL As SqlClient.SqlCommand
                    myConn.Open()
                    cmdSQL = New SqlClient.SqlCommand("sp_UpdateCLIDExpDate", myConn)
                    cmdSQL.CommandType = CommandType.StoredProcedure

                    cmdSQL.Parameters.Add("@Org", SqlDbType.VarChar, 20)
                    cmdSQL.Parameters("@Org").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@Org").Value = LoginData.OrgCode

                    cmdSQL.Parameters.Add("@ItemNO", SqlDbType.VarChar, 50)
                    cmdSQL.Parameters("@ItemNO").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@ItemNO").Value = ItemNO

                    cmdSQL.Parameters.Add("@LotNUM", SqlDbType.VarChar, 50)
                    cmdSQL.Parameters("@LotNUM").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@LotNUM").Value = LotNum

                    cmdSQL.Parameters.Add("@ExpDate", SqlDbType.DateTime)
                    cmdSQL.Parameters("@ExpDate").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@ExpDate").Value = ExpDate

                    cmdSQL.Parameters.Add("@user", SqlDbType.VarChar, 30)
                    cmdSQL.Parameters("@user").Direction = ParameterDirection.Input
                    cmdSQL.Parameters("@user").Value = LoginData.User

                    Dim x_status1 As String = cmdSQL.ExecuteScalar()
                    If x_status1 <> "1" Then
                        If x_status1 = "0" Then
                            x_msg = "update fail for eTrace"
                        Else
                            x_msg = "Invalid expiration date"
                        End If
                    End If
                    cmdSQL.Connection.Close()
                    'cmdSQL.Dispose()
                Else
                    x_msg = oCommand.Parameters("x_msg").Value.ToString
                End If
            Catch ex As Exception
                ErrorLogging("Update_LabelInfoHasSameLot", LoginData.User, ex.Message & ex.Source, "E")
                x_msg = ex.Message.ToString
            Finally
                If oCommand.Connection.State <> ConnectionState.Closed Then oCommand.Connection.Close()
                'oCommand.Dispose()
            End Try
            Return x_msg
        End Using
    End Function

    Public Function GetPalletList(ByVal PalletID As String) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getPallet", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5
            sda.SelectCommand.Parameters.Add("@Pallet", SqlDbType.VarChar, 50)
            sda.SelectCommand.Parameters("@Pallet").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@Pallet").Value = PalletID

            If FixNull(sda.SelectCommand.ExecuteScalar()).ToString = "0" Then
            Else
                sda.Fill(dsmodel)
            End If
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetPalletList", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function GetStatusList(ByVal TypeCode As String) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getStatusList", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5
            sda.SelectCommand.Parameters.Add("@Type_code", SqlDbType.VarChar, 50)
            sda.SelectCommand.Parameters("@Type_code").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@Type_code").Value = TypeCode

            sda.Fill(dsmodel)
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetStatusList", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function UpdatePalletList(ByVal ERPLogin As ERPLogin, ByVal myDashBoard As DashBoardData) As String

        Dim cmdSQL As SqlClient.SqlCommand
        Dim result As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_UpdatePallet", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@PalletID", SqlDbType.VarChar, 60)
            cmdSQL.Parameters("@PalletID").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@PalletID").Value = myDashBoard.PalletID

            cmdSQL.Parameters.Add("@Berth", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@Berth").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Berth").Value = myDashBoard.BerthID.ToString

            cmdSQL.Parameters.Add("@StatusDesc", SqlDbType.VarChar, 60)
            cmdSQL.Parameters("@StatusDesc").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@StatusDesc").Value = myDashBoard.StatusDesc

            cmdSQL.Parameters.Add("@User", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@User").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@User").Value = ERPLogin.User

            result = cmdSQL.ExecuteScalar()
            myConn.Close()
        Catch ex As Exception
            result = "003"
            ErrorLogging("UpdatePalletList", ERPLogin.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return result
    End Function

    Public Function GetMaterialInfo(ByVal RTNo As String, ByVal Material As String) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getMaterialInfo", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5

            sda.SelectCommand.Parameters.Add("@RTNo", SqlDbType.VarChar, 50)
            sda.SelectCommand.Parameters("@RTNo").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@RTNo").Value = RTNo

            sda.SelectCommand.Parameters.Add("@Material", SqlDbType.VarChar, 60)
            sda.SelectCommand.Parameters("@Material").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@Material").Value = Material

            sda.Fill(dsmodel)
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetMaterialInfo", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function UpdateMaterial(ByVal ERPLogin As ERPLogin, ByVal myDashBoard As DashBoardData) As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim result As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_UpdateMaterial", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@RTNO", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@RTNO").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@RTNO").Value = myDashBoard.RTNo

            cmdSQL.Parameters.Add("@Material", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@Material").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Material").Value = myDashBoard.Material

            cmdSQL.Parameters.Add("@Pallet", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@Pallet").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Pallet").Value = myDashBoard.PalletID

            cmdSQL.Parameters.Add("@StatusDesc", SqlDbType.VarChar, 60)
            cmdSQL.Parameters("@StatusDesc").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@StatusDesc").Value = myDashBoard.StatusDesc

            cmdSQL.Parameters.Add("@User", SqlDbType.VarChar, 40)
            cmdSQL.Parameters("@User").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@User").Value = ERPLogin.User

            result = cmdSQL.ExecuteScalar()
            myConn.Close()
        Catch ex As Exception
            result = "003"
            ErrorLogging("UpdatePalletList", ERPLogin.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return result
    End Function


    '---------------mic-----------------------------
    Public Function PLabelSEQ(ByVal mySEQ As LabelSEQ) As String
        Return ""
    End Function
    '---------------mic-----------------------------

#Region "SystemTables"
    Public Function Get_SysTableLists(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select * from T_SysTables ORDER BY Source, TableName")
                ds = da.ExecuteDataSet(Sqlstr, "SysTables")

                'Read eTraceDBConnString and back to Client
                ds.Tables(0).Columns.Add(New Data.DataColumn("DBConnString", System.Type.GetType("System.String")))
                ds.Tables(0).Rows(0)("DBConnString") = ConfigurationSettings.AppSettings.Item("eTraceDBConnString")

                Return ds

            Catch ex As Exception
                ErrorLogging("LabelBasicFunction-Get_SysTableLists", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    'Public Function Get_SysTableData(ByVal TableName As String, ByVal TableSource As String, ByVal LoginData As ERPLogin) As DataSet
    '    Using da As DataAccess = GetDataAccess()

    '        Dim dsItems As DataSet = New DataSet()

    '        Try

    '            If TableSource = "ETRACE" Then Return Get_eTraceTable(TableName, LoginData)

    '            'Column_Name	Data_Type	Nullable	Data_Default	Data_Length	   Primary_Key
    '            dsItems = Get_OraTableData(TableName, LoginData)
    '            dsItems.Tables(1).Columns.Add(New Data.DataColumn("Primary_Key", System.Type.GetType("System.String")))

    '            Return dsItems

    '        Catch oe As Exception
    '            ErrorLogging("LabelBasicFunction-Get_SysTableData", LoginData.User.ToUpper, "TableName: " & TableName & "; " & oe.Message & oe.Source, "E")
    '            Return Nothing
    '        End Try
    '    End Using

    'End Function

    'Public Function Get_eTraceTable(ByVal TableName As String, ByVal LoginData As ERPLogin) As DataSet
    '    Using da As DataAccess = GetDataAccess()

    '        Dim dsItems As DataSet = New DataSet()

    '        Try
    '            Dim Sqlstr, TableSql As String
    '            TableSql = "Select * from " & TableName
    '            Sqlstr = String.Format(TableSql)
    '            'dsItems = da.ExecuteDataSet(Sqlstr, "tabledata")

    '            Dim dtData As New DataTable()
    '            dtData = da.GetDataSchema(Sqlstr)
    '            dtData.TableName = "tabledata"
    '            dsItems.Tables.Add(dtData)


    '            Dim ds As DataSet = New DataSet()
    '            'Sqlstr = String.Format("Select a.name AS Column_Name, b.name AS Data_Type, CASE WHEN a.is_nullable = 1 THEN 'Y' ELSE 'N' END AS Nullable, ISNULL(d.text, '') " _
    '            '       & " AS Data_Default, CASE WHEN ColumnProperty(a.object_id, a.Name, 'Precision') = - 1 THEN '2^31-1' ELSE rtrim(ColumnProperty(a.object_id, a.Name, " _
    '            '       & " 'Precision')) END AS Data_Length, CASE WHEN EXISTS (SELECT 1 FROM sys.objects x JOIN sys.indexes y ON x.Type = N'PK' AND x.Name = y.Name JOIN " _
    '            '       & " sysindexkeys z ON z.ID = a.Object_id AND z.indid = y.index_id AND z.Colid = a.Column_id) THEN 'Y' ELSE '' END AS Primary_Key FROM sys.columns AS a " _
    '            '       & " LEFT OUTER JOIN sys.types AS b ON a.user_type_id = b.user_type_id INNER JOIN sys.objects AS c ON a.object_id = c.object_id AND c.type = 'U' " _
    '            '       & " LEFT OUTER JOIN sys.syscomments AS d ON a.default_object_id = d.id LEFT OUTER JOIN sys.extended_properties AS e ON e.major_id = c.object_id " _
    '            '       & " AND e.minor_id = a.column_id AND e.class = 1 LEFT OUTER JOIN sys.extended_properties AS f ON f.major_id = c.object_id AND f.minor_id = 0 AND f.class = 1 " _
    '            '       & " WHERE (c.is_ms_shipped = 0) AND (c.name = '{0}' ) ORDER BY c.name, a.column_id", TableName)
    '            Sqlstr = String.Format("exec sp_GetTableStructure '{0}'", TableName)
    '            ds = da.ExecuteDataSet(Sqlstr, "structure")

    '            dsItems.Merge(ds.Tables(0))
    '            Return dsItems

    '        Catch oe As Exception
    '            ErrorLogging("LabelBasicFunction-Get_eTraceTable", LoginData.User.ToUpper, "TableName: " & TableName & "; " & oe.Message & oe.Source, "E")
    '            Return Nothing
    '        End Try
    '    End Using

    'End Function

    Public Function Get_OraTableData(ByVal TableName As String, ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()


            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_oe_pkg.get_tb_info"              'Get Customized Table Lists
                oda.SelectCommand.Parameters.Add("p_tb_name", OracleType.VarChar, 100).Value = TableName
                oda.SelectCommand.Parameters.Add("o_tb_data", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_tb_structure", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "tabledata"
                ds.Tables(1).TableName = "structure"

                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_OraTableData", LoginData.User.ToUpper, "TableName: " & TableName & "; " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    'Public Function Save_SysTableData(ByVal TableName As String, ByVal TableSource As String, ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String
    '    Using da As DataAccess = GetDataAccess()

    '        Dim SaveFlag, ErrMsg As String
    '        ErrMsg = "Update Table " & TableName & " error "


    '        If TableSource = "ETRACE" Then
    '            Try
    '                Dim mydtData As New DataTable
    '                mydtData = Items.Tables("tabledata")
    '                SaveFlag = da.UpdateSysTable(TableName, mydtData)

    '            Catch ex As Exception
    '                SaveFlag = ErrMsg & ": " & ex.Message & ex.Source
    '                ErrorLogging("LabelBasicFunction-Save_SysTableData", LoginData.User.ToUpper, SaveFlag)
    '            End Try

    '            Return SaveFlag
    '        End If


    '        Try

    '            'Delete original data from Custom Table before insert new data
    '            If Items.Tables("data").Rows.Count > 0 Then
    '                Dim DelFlag As String
    '                DelFlag = Delete_SysTableData(TableName, LoginData)
    '                If DelFlag <> "Y" Then Return DelFlag
    '            End If

    '            'If no new Data needs to insert, then return to client
    '            Items.Tables("tabledata").AcceptChanges()
    '            If Items.Tables("tabledata").Rows.Count = 0 Then Return "Y"

    '            'Insert new data for the first time
    '            SaveFlag = ProcessSysTables(TableName, True, Items, LoginData)

    '            'If Data saved failed, then insert the original data for recovery if there exists
    '            If SaveFlag <> "Y" AndAlso Items.Tables("data").Rows.Count > 0 Then
    '                ProcessSysTables(TableName, False, Items, LoginData)
    '            End If
    '            Return SaveFlag

    '        Catch ex As Exception
    '            ErrorLogging("LabelBasicFunction-Save_SysTableData", LoginData.User.ToUpper, "TableName: " & TableName & "; " & ex.Message & ex.Source, "E")
    '            Return ex.Message
    '        End Try

    '    End Using
    'End Function

    Public Function ProcessSysTables(ByVal TableName As String, ByVal NewData As Boolean, ByVal Items As DataSet, ByVal LoginData As ERPLogin) As String

        ProcessSysTables = ""

        Dim i, j, k As Integer
        Dim dtData, dtStructure As New DataTable

        Try

            If NewData = True Then
                dtData = Items.Tables("tabledata").Copy
                dtStructure = Items.Tables("structure").Copy
            Else
                dtData = Items.Tables("data").Copy
                dtStructure = Items.Tables("structure").Copy
            End If

            For i = 0 To dtData.Rows.Count - 1
                Dim DataType, DataStr, StrSql As String
                DataType = ""
                DataStr = ""
                StrSql = ""

                Dim CountNull As Integer = 0
                For j = 0 To dtData.Columns.Count - 1
                    If dtData.Rows(i)(j) Is DBNull.Value Then
                        DataStr = "NULL"
                        CountNull = CountNull + 1
                    Else
                        For k = j To dtStructure.Rows.Count - 1
                            If k = j Then
                                DataType = dtStructure.Rows(k)("DATA_TYPE").ToString.ToUpper
                                If DataType.Contains("VARCH") Then DataType = "VARCH"
                                Exit For
                            End If
                        Next

                        If DataType = "VARCH" Then
                            If dtData.Rows(i)(j) = "" Then
                                DataStr = "NULL"
                                CountNull = CountNull + 1
                            Else
                                DataStr = DataStr.Replace("'", ".")
                                DataStr = "'" & dtData.Rows(i)(j) & "'"
                                'DataStr = "''" & dtData.Rows(i)(j) & "''"
                            End If
                        ElseIf DataType = "NUMBER" Then
                            DataStr = dtData.Rows(i)(j).ToString
                        ElseIf DataType = "DATE" Then
                            Dim ValDate As Date = CDate(dtData.Rows(i)(j))
                            DataStr = ValDate.ToString("yyyy/MM/dd")
                            DataStr = "apps.fnd_date.canonical_to_date('" & DataStr & "')"
                        End If
                    End If

                    If j = dtData.Columns.Count - 1 Then
                        StrSql = StrSql & DataStr
                    Else
                        StrSql = StrSql & DataStr & ","
                    End If
                Next

                'If all columns with value = NULL, then we will ignore this row
                If CountNull = dtStructure.Rows.Count Then Continue For

                'Insert Data to Oracle Custom Table one row at one time
                ProcessSysTables = Insert_DataToTable(TableName, StrSql, LoginData)
                If ProcessSysTables <> "Y" Then
                    ProcessSysTables = "Error Row: " & i + 1 & "; " & ProcessSysTables
                    ErrorLogging("LabelBasicFunction-ProcessSysTables", LoginData.User.ToUpper, "TableName: " & TableName & "; " & ProcessSysTables)
                    Exit For
                End If

            Next
            Return ProcessSysTables

        Catch ex As Exception
            ErrorLogging("LabelBasicFunction-ProcessSysTables", LoginData.User.ToUpper, "TableName: " & TableName & "; " & ex.Message & ex.Source, "E")
            Return ex.Message
        End Try

    End Function

    Public Function Insert_DataToTable(ByVal TableName As String, ByVal StrSql As String, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_oe_pkg.insert_datatotable"
                OC.Parameters.Add("p_tb_name", OracleType.VarChar, 100).Value = TableName
                OC.Parameters.Add("p_strsql", OracleType.VarChar, 1000).Value = StrSql
                OC.Parameters.Add("o_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_message", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())
                Dim Flag, Msg As String
                Flag = OC.Parameters("o_flag").Value.ToString
                Msg = OC.Parameters("o_message").Value.ToString
                OC.Connection.Close()

                If Flag = "Y" Then
                    Insert_DataToTable = Flag
                Else
                    Insert_DataToTable = Msg
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Insert_DataToTable", LoginData.User.ToUpper, "TableName: " & TableName & "; " & oe.Message & oe.Source, "E")
                Insert_DataToTable = oe.Message
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Delete_SysTableData(ByVal TableName As String, ByVal LoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()
            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_oe_pkg.delete_custom_table"
                OC.Parameters.Add("p_tb_name", OracleType.VarChar, 100).Value = TableName
                OC.Parameters.Add("o_flag", OracleType.VarChar, 20).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_message", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())
                Dim Flag, Msg As String
                Flag = OC.Parameters("o_flag").Value.ToString
                Msg = OC.Parameters("o_message").Value.ToString
                OC.Connection.Close()

                If Flag = "Y" Then
                    Delete_SysTableData = Flag
                Else
                    Delete_SysTableData = Msg
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_SysTableData", LoginData.User.ToUpper, "TableName: " & TableName & "; " & oe.Message & oe.Source, "E")
                Delete_SysTableData = oe.Message
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

#End Region

End Class

