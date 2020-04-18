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

Public Structure GRHeaderStructure
    Public Type As Integer
    Public InvoiceNo As String                                                     ' InvoiceNo
    Public BillOfLading As String
    Public DeliveryNote As String
    Public PostDate As Date
    Public HeaderText As String                                                  ' TruckNo
    Public OrderNo As String
    Public OrderItem As String
    Public CreatedBy As String
    Public CreatedOn As Date
    Public ChangedBy As String
    Public ChangedOn As Date
    Public GRNo As String
    Public POCurrency As String
    Public OrderType As Integer
    Public AllowAMLUpdate As String
    Public CLIDFlag As Boolean
    Public POLists As DataSet
End Structure

Public Structure CreateGRResponse
    Public GRNo As String
    Public PostDate As Date
    Public GRMessage As String
    Public GRStatus As String
    Public CLIDs As DataSet
    Public GRMsg As DataSet
End Structure

Public Structure MatLabel
    Public LabelID As String
    Public OrgCode As String
    Public RT As String
    Public Material As String
    Public MatSuffix As String
    Public Rev As String
    Public Qty As String
    Public UOM As String
    Public ExpDate As String
    Public RoHS As String
    Public DCode As String
    Public LotNo As String
    Public InspFlag As String
    Public RecDate As String
    Public LabelType As String
    Public Safety As String
    Public ESD As String
    Public Stemp As String
    Public MSL As String
    Public Mfr As String
    Public MfrPN As String
    Public MCPosition As String
    Public PreSubInv As String
    Public ItemText As String
    Public COO As String
    Public PNSeq As String
    Public IPN As String
End Structure

Public Structure RTLabel
    Public RT As String
    Public RecDate As String
    Public Priority As String
    Public DestSubInv As String
    Public DestLocator As String
    Public StorageType As String
    Public MCPosition As String
    Public OrgCode As String
    Public OrgName As String
    Public Vendor As String
    Public VendorName As String
    Public InvoiceDN As String
    Public PONo As String
    Public POLine As String
    Public Shipment As String
    Public POType As String
    Public VMIFlag As String
    Public Buyer As String
    Public BuyerName As String
    Public Material As String
    Public MatDesc As String
    Public Qty As String
    Public UOM As String
    Public NoOfPackage As String
    Public AddlText As String
    Public SysRev As String
    Public PORev As String
    Public RoHS As String
    Public ExpDate As String
    Public InspFlag As String
    Public FAIStatus As String
    Public TempLoc As String
    Public ConFactor As String
    Public Safety As String
    Public Size As String
    Public MSL As String
    Public Stemp As String
    Public COC As String
    Public TransactBy As String
    Public OptrName As String
    Public Mfr As String
    Public MfrPN As String
    Public QMLStatus As String
    Public DateCode As String
    Public LotNo As String
    Public ESD As String
    Public ShelfLife As String
    Public HisInsRT As String
    Public SampleLoc As String
End Structure

Public Structure MGTraceData
    Public TraceabilityLevel As String
    Public AddlData As String
End Structure

Public Structure POData               'Split PO Data
    Public PONo As String
    Public ReleaseNo As String
    Public LineNo As String
    Public ShipmentNo As String
    Public Distribution As String
    Public MaterialNo As String
    Public Locator As String
End Structure

Public Structure IRData
    Public User As String
    Public SourceOrg As String
    Public DestOrg As String
    Public ShipmentNo As String
    Public IRRTNo As String
End Structure

Public Structure RcvID
    Public GroupID As Integer
    Public HeaderID As Integer
End Structure


Public Class Receiving            ' Publish target location:  C:\Inetpub\wwwroot\eTrace_OracleERP
    Inherits PublicFunction
    Private RTLabelFile As String = "D:\eTrace\RTLabel.lab"  'For Normal Printer, page format is A4 
    Private CLIDLabelFile As String = "D:\eTrace\CLMatLabel.lab"  'For zebra Printer QL420
    Private ESDImage As String = "D:\eTrace\ESD.jpg"
    Private BlankESD As String = "D:\eTrace\BlankESD.jpg"
    Public Function GetInspectionRT(ByVal OrgCode As String, ByVal Item As String, ByVal VendorID As String, ByVal DateCode As String, ByVal LotNo As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim ds As DataSet = New DataSet
            Dim RT As String
            Dim sda As SqlClient.SqlDataAdapter
            Try
                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_GetInspectionRT"
                sda.SelectCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20).Value = OrgCode
                sda.SelectCommand.Parameters.Add("@Item", SqlDbType.VarChar, 50).Value = Item
                sda.SelectCommand.Parameters.Add("@VendorID", SqlDbType.VarChar, 50).Value = VendorID
                sda.SelectCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 50).Value = DateCode
                sda.SelectCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 50).Value = LotNo
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("Receiving-GetInspectionRT", "", "OrgCode: " & OrgCode & ", " & "Item: " & Item & ", " & "VendorID: " & VendorID & ", " & ex.Message & ex.Source, "E")
            End Try
            If ds.Tables(0).Rows.Count > 0 Then
                RT = ds.Tables(0).Rows(0)(0)
            Else
                RT = ""
            End If
            Return RT
        End Using
    End Function
    Public Function CheckItemMPN(ByVal Item As String, ByVal ManufacturerPN As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim Sqlstr As String
            Dim SuppData As String
            Try
                Sqlstr = String.Format("select top 1 SupplierData from T_RecMPN where ItemNo = '{0}' and ManufacturerPN = '{1}' ", Item, ManufacturerPN)
                SuppData = Convert.ToString(da.ExecuteScalar(Sqlstr))
                If SuppData <> "" Then
                    CheckItemMPN = SuppData
                Else
                    CheckItemMPN = ""
                End If
            Catch ex As Exception
                ErrorLogging("Receiving-CheckItemMPN", "", ex.Message & ex.Source, "E")
                CheckItemMPN = ""
            End Try
        End Using
    End Function
    Public Function GetPOLineMPQ(ByVal p_ds As DataTable) As DataTable
        Using da As DataAccess = GetDataAccess()
            Dim Oda As OracleDataAdapter = da.Oda_Insert()
            Dim i As Integer

            For i = 0 To p_ds.Rows.Count - 1
                If p_ds.Rows(i).RowState <> DataRowState.Added Then
                    p_ds.Rows(i).SetAdded()
                End If
            Next
            Try
                Oda.InsertCommand.CommandType = CommandType.StoredProcedure
                Oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.get_item_mpq"
                Oda.InsertCommand.Parameters.Add("p_org_id", OracleType.Int32)
                Oda.InsertCommand.Parameters.Add("p_po_line_id", OracleType.Int32)
                Oda.InsertCommand.Parameters.Add("o_item_mpq", OracleType.Double)
                Oda.InsertCommand.Parameters("o_item_mpq").Direction = ParameterDirection.Output
                Oda.InsertCommand.Parameters("p_org_id").SourceColumn = "SHIP_TO_ORGANIZATION_ID"
                Oda.InsertCommand.Parameters("p_po_line_id").SourceColumn = "PO_LINE_ID"
                Oda.InsertCommand.Parameters("o_item_mpq").SourceColumn = "MPQ"
                Oda.InsertCommand.Connection.Open()
                Oda.Update(p_ds)
                Oda.InsertCommand.Connection.Close()
                Return p_ds
            Catch oe As Exception
                ErrorLogging("Receiving-GetPOLineMPQ", "", oe.Message & oe.Source, "E")
                Return p_ds
            Finally
                If Oda.InsertCommand.Connection.State <> ConnectionState.Closed Then Oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function GetItemMPQ(ByVal p_org_id As Int32, ByVal p_po_line_id As Integer) As Double
        Using da As DataAccess = GetDataAccess()
            Dim MPQ As Double
            Dim Oda As OracleCommand = da.OraCommand()
            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.get_item_mpq"
                Oda.Parameters.Add("p_org_id", OracleType.Int32).Value = p_org_id
                Oda.Parameters.Add("p_po_line_id", OracleType.Int32).Value = p_po_line_id
                Oda.Parameters.Add("o_item_mpq", OracleType.Double)
                Oda.Parameters("o_item_mpq").Direction = ParameterDirection.Output
                Oda.Connection.Open()
                Oda.ExecuteNonQuery()
                MPQ = Oda.Parameters("o_item_mpq").Value
                Oda.Connection.Close()
                Return MPQ
            Catch oe As Exception
                ErrorLogging("Receiving-GetItemMPQ", "", "OrgID: " & p_org_id & ", " & "POLineID: " & p_po_line_id & ", " & oe.Message & oe.Source, "E")
                Return 0
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetRecData(ByVal eTraceModule As String, ByVal TransactionID As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim LoginData As DataSet = New DataSet

            Try
                LoginData = GetLoginData(eTraceModule, TransactionID)

                'Read problematic D/C & L/C from eTrace table    02/16/2016
                Dim Sqlstr As String
                Dim dsDCLN As DataSet = New DataSet
                Sqlstr = String.Format("Select * from T_DateCodeLotNo ")
                dsDCLN = da.ExecuteDataSet(Sqlstr, "DateCodeLN")
                LoginData.Merge(dsDCLN.Tables(0))

                Dim ds As DataSet = New DataSet
                Sqlstr = String.Format("Select ProcessID as PrCLID from T_SysLOV with (nolock) where Name = 'REC Black list for CLID upload'  ")
                ds = da.ExecuteDataSet(Sqlstr, "BLKCLID")
                LoginData.Merge(ds.Tables(0))

                Return LoginData

            Catch ex As Exception
                ErrorLogging("Receiving-GetRecData", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using
    End Function


    Public Function GetItems(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet

        Try
            If GRHeader.Type = 8 Or GRHeader.Type = 10 Then
                Return ReadGRItems(LoginData, GRHeader)
            ElseIf GRHeader.Type = 11 Then
                Return VerifyRTData(LoginData, GRHeader)
            ElseIf GRHeader.Type = 12 Then
                Return ValidateDHLData(LoginData, GRHeader)
            End If


            Dim ERPItems As DataSet = New DataSet
            ERPItems = GetItemsFromERP(LoginData, GRHeader)
            If GRHeader.Type <> 9 Then Return ERPItems

            ' CLID creation based on Oracle Receipt, so we read RT data from Oracle first and also check if the Item/RT already exists in eTrace
            If ERPItems.Tables("ErrorTable").Rows.Count > 0 Then
                Return ERPItems
                Exit Function
            End If

            Dim myCLIDs As DataSet = New DataSet
            myCLIDs = ReadGRItems(LoginData, GRHeader)

            Dim i, j As Integer
            Dim DR() As DataRow = Nothing

            If myCLIDs.Tables("TraceabilityData").Rows.Count > 0 Then
                For i = 0 To ERPItems.Tables("POItemData").Rows.Count - 1
                    Dim PONo, POItem, POLineID As String
                    PONo = ERPItems.Tables("POItemData").Rows(i)("PONo").ToString
                    POItem = ERPItems.Tables("POItemData").Rows(i)("POItem").ToString
                    POLineID = ERPItems.Tables("POItemData").Rows(i)("POLineID").ToString       'TransactionID

                    Dim RTQty As Decimal = 0
                    'DR = myCLIDs.Tables("TraceabilityData").Select(" PONo = '" & PONo & "' and POItem = '" & POItem & "' and Status = '" & POLineID & "'")
                    DR = myCLIDs.Tables("TraceabilityData").Select(" PONo = '" & PONo & "' and POItem = '" & POItem & "' and POLineID = '" & POLineID & "'")
                    If DR.Length > 0 Then
                        For j = 0 To DR.Length - 1
                            RTQty = RTQty + DR(j)("Qty")
                        Next

                        RTQty = ERPItems.Tables("POItemData").Rows(i)("QtyBaseUOM") - RTQty
                        ERPItems.Tables("POItemData").Rows(i)("QtyBaseUOM") = RTQty
                        ERPItems.Tables("POItemData").Rows(i)("GRQty") = RTQty
                        ERPItems.Tables("POItemData").Rows(i).AcceptChanges()

                    End If
                Next
                ERPItems.Tables("POItemData").AcceptChanges()

                Dim Qty As Decimal = 0
                Dim sqlstr As String = " QtyBaseUOM <= '" & Qty & "'"             'Find zero Qty items
                DR = Nothing
                DR = ERPItems.Tables("POItemData").Select(sqlstr)
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    ERPItems.Tables("POItemData").AcceptChanges()
                End If
            End If

            If ERPItems.Tables("POItemData").Rows.Count = 0 Then
                Dim myDataRow As Data.DataRow
                myDataRow = ERPItems.Tables("ErrorTable").NewRow
                myDataRow("ErrorMsg") = "^REC-74@"
                ERPItems.Tables("ErrorTable").Rows.Add(myDataRow)
            End If

            Return ERPItems

        Catch ex As Exception
            ErrorLogging("Receiving-GetItems", LoginData.User.ToUpper, "OrderNo: " & GRHeader.OrderNo & " with Receipt type " & GRHeader.Type & ", " & ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Public Function VerifyRTData(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("exec sp_VerifyRTData '{0}', '{1}' ", LoginData.OrgCode, GRHeader.OrderNo)
                Dim sql() As String = {Sqlstr}
                Dim tables() As String = {"dtMsg", "CLIDs"}
                Return da.ExecuteDataSet(sql, tables)

            Catch ex As Exception
                ErrorLogging("Receiving-VerifyRTData", LoginData.User.ToUpper, "RTNo: " & GRHeader.OrderNo & "; " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Private Function VerifyRTCLID(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Dim CLID As String = Header.OrderItem.ToString

            Try
                Dim Sqlstr As String
                Dim RtnMsg As String = ""
                Sqlstr = String.Format("exec sp_VerifyRTCLID '{0}', N'{1}' ", CLID, LoginData.User)
                RtnMsg = da.ExecuteScalar(Sqlstr).ToString

                VerifyRTCLID.GRStatus = RtnMsg
                If RtnMsg = "Y" Then
                    VerifyRTCLID.GRMessage = "CLID updated successfully "
                Else
                    VerifyRTCLID.GRMessage = "CLID updated failed "
                End If

            Catch ex As Exception
                ErrorLogging("Receiving-VerifyRTCLID", LoginData.User.ToUpper, "RTNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
            End Try

        End Using

    End Function

    Public Function ValidateDHLData(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim TruckNo As String = GRHeader.OrderNo

            Try
                Dim Sqlstr As String
                Dim ds As New DataSet
                Dim CheckFlag As String = ""

                'Donwload Truck Data here
                If TruckNo <> "" AndAlso GRHeader.POLists Is Nothing Then
                    CheckFlag = "N"
                    Sqlstr = String.Format("exec sp_RECCheckTruck '{0}', '{1}' ", TruckNo, CheckFlag)
                    ds = da.ExecuteDataSet(Sqlstr, "DHLData")
                    Return ds
                End If


                If Not GRHeader.POLists Is Nothing Then

                    Dim dtTruck As DataTable
                    dtTruck = New Data.DataTable("dtTruck")
                    dtTruck.Columns.Add(New Data.DataColumn("TruckID", System.Type.GetType("System.String")))
                    dtTruck.Columns.Add(New Data.DataColumn("Flag", System.Type.GetType("System.String")))

                    Dim i, j As Integer
                    Dim DR() As DataRow
                    Dim myDR As Data.DataRow

                    Dim POLists As DataSet = Nothing
                    POLists = New DataSet
                    POLists = GRHeader.POLists

                    j = 0
                    For i = 0 To POLists.Tables(0).Rows.Count - 1
                        If POLists.Tables(0).Rows(i)("TruckID") Is DBNull.Value Then
                        ElseIf POLists.Tables(0).Rows(i)("TruckID") = "" Then
                        Else
                            TruckNo = POLists.Tables(0).Rows(i)("TruckID")

                            j = j + 1
                            POLists.Tables(0).Rows(i)("LineNo") = j

                            DR = Nothing
                            DR = dtTruck.Select(" TruckID = '" & TruckNo & "'")
                            If DR.Length = 0 Then
                                Dim myResult As String = "N"

                                CheckFlag = "Y"
                                Sqlstr = String.Format("exec sp_RECCheckTruck '{0}', '{1}' ", TruckNo, CheckFlag)
                                ds = da.ExecuteDataSet(Sqlstr, "DHLData")
                                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                                ElseIf ds.Tables(0).Rows.Count > 0 Then
                                    myResult = "Y"
                                    POLists.Tables(0).Rows(i)("Flag") = myResult
                                End If

                                myDR = dtTruck.NewRow()
                                myDR("TruckID") = TruckNo
                                myDR("Flag") = myResult
                                dtTruck.Rows.Add(myDR)
                            End If
                            POLists.AcceptChanges()
                        End If
                    Next

                    Return POLists
                End If

            Catch ex As Exception
                ErrorLogging("Receiving-ValidateDHLData", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Private Function SaveDHLData(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Try
                Dim Sqlstr As String
                Dim RtnMsg As String = ""

                Items.DataSetName = "dsItem"
                Sqlstr = String.Format("exec sp_RECSaveTruck '{0}', N'{1}' ", LoginData.User, DStoXML(Items))
                RtnMsg = da.ExecuteScalar(Sqlstr).ToString

                SaveDHLData.GRStatus = RtnMsg
                If RtnMsg = "Y" Then
                    SaveDHLData.GRMessage = "DHL Truck Data saved successfully "
                Else
                    SaveDHLData.GRMessage = "DHL Truck Data saved failed "
                End If

            Catch ex As Exception
                ErrorLogging("Receiving-SaveDHLData", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try

        End Using
    End Function
    Public Function DownloadSuppData(ByVal ItemNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As DataSet = New DataSet
            Dim sda As SqlClient.SqlDataAdapter
            sda = da.Sda_Sele()
            Try
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "sp_RECDownloadSuppData"
                sda.SelectCommand.Parameters.Add("@ItemNo", SqlDbType.VarChar, 150).Value = ItemNo
                sda.SelectCommand.CommandTimeout = TimeOut_M5
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()
            Catch ex As Exception
                ErrorLogging("Receiving-DownloadSuppData", "", "ItemNo: " & ItemNo & ", " & ex.Message & ex.Source, "E")
            Finally
                If sda.SelectCommand.Connection.State <> ConnectionState.Closed Then sda.SelectCommand.Connection.Close()
            End Try
            Return ds
        End Using
    End Function
    Public Function SaveSuppData(ByVal LoginData As ERPLogin, ByVal Items As DataSet) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()
            Dim sda As SqlClient.SqlDataAdapter
            Dim Sqlstr As String
            sda = da.Sda_Insert()
            Try
                Sqlstr = String.Format("Delete T_RecMPN ")
                da.ExecuteNonQuery(Sqlstr)

                Dim i As Integer
                For i = 0 To Items.Tables(0).Rows.Count - 1
                    If Items.Tables(0).Rows(i).RowState <> DataRowState.Added Then
                        Items.Tables(0).Rows(i).SetAdded()
                    End If
                Next
                sda.InsertCommand.CommandType = CommandType.StoredProcedure
                sda.InsertCommand.CommandText = "sp_RECSaveSuppData"
                sda.InsertCommand.CommandTimeout = TimeOut_M5
                sda.InsertCommand.Parameters.Add("@User", SqlDbType.NVarChar, 50)
                sda.InsertCommand.Parameters.Add("@ItemNo", SqlDbType.VarChar, 150)
                sda.InsertCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 150)
                sda.InsertCommand.Parameters.Add("@SupplierData", SqlDbType.VarChar, 300)
                sda.InsertCommand.Parameters.Add("@Flag", SqlDbType.VarChar, 1).Direction = ParameterDirection.Output
                sda.InsertCommand.Parameters("@User").Value = LoginData.User.ToUpper
                sda.InsertCommand.Parameters("@ItemNo").SourceColumn = "ItemNumber"
                sda.InsertCommand.Parameters("@ManufacturerPN").SourceColumn = "ManufacturerPN"
                sda.InsertCommand.Parameters("@SupplierData").SourceColumn = "SupplierData"
                sda.InsertCommand.Parameters("@Flag").SourceColumn = "Flag"
                sda.InsertCommand.Connection.Open()
                sda.Update(Items.Tables(0))
                sda.InsertCommand.Connection.Close()
                Dim dr() As DataRow = Nothing
                dr = Items.Tables(0).Select("Flag = 'N'")
                If dr.Length > 0 Then
                    SaveSuppData.GRStatus = "N"
                    SaveSuppData.GRMessage = "Supplier Data saved failed "
                Else
                    SaveSuppData.GRStatus = "Y"
                    SaveSuppData.GRMessage = "Supplier Data saved successfully "
                End If
            Catch ex As Exception
                ErrorLogging("Receiving-SaveSuppData", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Finally
                If sda.InsertCommand.Connection.State <> ConnectionState.Closed Then sda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function
    Public Function AddGRTables() As DataSet

        Dim GRItems As DataSet = New DataSet
        Dim POItemData As DataTable
        Dim myDataColumn As DataColumn

        POItemData = New Data.DataTable("POItemData")
        myDataColumn = New Data.DataColumn("PONo", System.Type.GetType("System.String"))                 '0
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("POItem", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialRevision", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MaterialDesc", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Decimal"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BaseUOM", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("EntryQty", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("NeedbyDate", System.Type.GetType("System.DateTime"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PromisedDate", System.Type.GetType("System.DateTime"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ItemText", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubInventory", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StoragePosition", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("DeliveryType", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Operator", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ReasonCode", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ProductionDate", System.Type.GetType("System.DateTime"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.DateTime"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("VendorNo", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("VendorName", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("VendorPart", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("VendorPartScan", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MatSuffix1", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MatSuffix2", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MatSuffix3", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BoxDesc", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("QMLStatus", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("UOM", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BaseUOMConv", System.Type.GetType("System.Decimal"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RoHS", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StatusCode", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SLED", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StockType", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Traceable", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("AddlData", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("GRQty", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExpDays", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MinRemDays", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("EarlyDays", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PredefinedSubInv", System.Type.GetType("System.String"))         'PredefinedSubInv,PredefinedLocator
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PredefinedLocator", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("POLineID", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PullMethod", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RevControl", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LotControl", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExpControl", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Stemp", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MSL", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("COC", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("EER", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RESData", System.Type.GetType("System.String"))
        POItemData.Columns.Add(myDataColumn)

        GRItems.Tables.Add(POItemData)

        Dim TraceabilityData As DataTable
        TraceabilityData = New Data.DataTable("TraceabilityData")
        myDataColumn = New Data.DataColumn("POItem", System.Type.GetType("System.String"))        '0
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("DateCode", System.Type.GetType("System.String"))      '1
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LotNo", System.Type.GetType("System.String"))         '2
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("COO", System.Type.GetType("System.String"))           '3
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Qty", System.Type.GetType("System.String"))           '4
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LID", System.Type.GetType("System.String"))           '5
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PONo", System.Type.GetType("System.String"))          '6
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MatNo", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RefCLID", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("BoxID", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubInventory", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PredefinedSubInv", System.Type.GetType("System.String"))         'PredefinedSubInv,PredefinedLocator
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("PredefinedLocator", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("POLineID", System.Type.GetType("System.String"))
        TraceabilityData.Columns.Add(myDataColumn)

        GRItems.Tables.Add(TraceabilityData)

        Dim ErrorTable As DataTable
        ErrorTable = New Data.DataTable("ErrorTable")
        myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
        ErrorTable.Columns.Add(myDataColumn)

        GRItems.Tables.Add(ErrorTable)

        Return GRItems

    End Function

    Private Function ReadGRItems(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet
        Using da As DataAccess = GetDataAccess()

            ReadGRItems = New DataSet
            ReadGRItems = AddGRTables()

            Dim POItemData As DataTable
            Dim TraceabilityData As DataTable
            Dim ErrorTable As DataTable
            Dim PoItemRow As Data.DataRow
            Dim TraceabilityDataRow As Data.DataRow
            Dim ErrorRow As Data.DataRow

            POItemData = ReadGRItems.Tables(0).Clone
            TraceabilityData = ReadGRItems.Tables(1).Clone
            ErrorTable = ReadGRItems.Tables(2).Clone

            ReadGRItems = New DataSet
            ReadGRItems.Tables.Add(POItemData)
            ReadGRItems.Tables.Add(TraceabilityData)
            ReadGRItems.Tables.Add(ErrorTable)

            Dim i, j As Integer

            Dim myCLIDs As DataSet = New DataSet
            Dim ProcessID As Integer

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select InvoiceNo,BillOfLading,DN,RecDate,HeaderText,CreatedBy,CreatedOn,ChangedBy,ChangedOn,RecDocItem,MaterialNo,MaterialRevision, MaterialDesc,BaseUOM,VendorPN,VendorID,VendorName,StoragePosition,ProdDate,DeliveryType,Operator,ReasonCode,AddlText,ExpDate,MatSuffix1,MatSuffix2,MatSuffix3,DateCode,LotNo,CountryOfOrigin,QtyBaseUOM,CLID,PurOrdNo,PurOrdItem,IsTraceable,StatusCode,ReferenceCLID,BoxID,Manufacturer,ManufacturerPN,QMLStatus,AddlData,PredefinedSubInv,PredefinedLocator,ProcessID,MatDocNo,MatDocItem,SLOC,StorageBin, RTLot, Status ='' from T_CLMaster with (nolock) where RecDocNo='{0}' and OrgCode='{1}' ", GRHeader.OrderNo, LoginData.OrgCode)
                myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 OrElse myCLIDs.Tables(0).Rows.Count = 0 Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrorMsg") = "^REC-31@"      'Error Message: Receipt Number not found or invalid
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                ProcessID = IIf(myCLIDs.Tables(0).Rows(0)("ProcessID") Is DBNull.Value, 0, myCLIDs.Tables(0).Rows(0)("ProcessID"))

                'If GRHeader.Type = 8 AndAlso (ProcessID = 401 OrElse ProcessID = 402 OrElse ProcessID = 403) Then
                If GRHeader.Type = 8 AndAlso Left(ProcessID.ToString, 1) = "4" Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrorMsg") = "^REC-89@"      'Error Message: IR can't be resersed. Please ask vendor to create return IR to continue
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                Dim DR() As DataRow = Nothing
                Dim ds As DataSet = New DataSet
                ds = myCLIDs.Copy

                If GRHeader.Type = 8 OrElse GRHeader.Type = 10 Then
                    If GRHeader.Type = 8 Then
                        DR = myCLIDs.Tables(0).Select(" StatusCode = '1' or StatusCode = '9' ")
                        If DR.Length = 0 Then
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow("ErrorMsg") = "^REC-31@"
                            ErrorTable.Rows.Add(ErrorRow)
                            Exit Function
                        End If
                    End If

                    DR = Nothing
                    DR = myCLIDs.Tables(0).Select(" StatusCode = '0' and RecDocItem <> '' and ReferenceCLID IS NULL ")      ' Set StatusCode = '0' for disabled CLIDs
                    If DR.Length > 0 Then
                        For j = 0 To DR.Length - 1
                            DR(j)("Status") = "X"
                            DR(j).AcceptChanges()
                        Next
                        myCLIDs.Tables(0).AcceptChanges()
                    End If

                    DR = Nothing
                    DR = ds.Tables(0).Select(" StatusCode = '0' and RecDocItem IS NULL ")
                    If DR.Length > 0 Then
                        For i = 0 To DR.Length - 1
                            Dim CLID As String = DR(i)("CLID").ToString
                            Dim CLIDRow() As DataRow = Nothing
                            CLIDRow = myCLIDs.Tables(0).Select(" StatusCode = '0' and ReferenceCLID = '" & CLID & "'")      ' Set StatusCode = '0' for merged CLIDs
                            If CLIDRow.Length > 0 Then
                                For j = 0 To CLIDRow.Length - 1
                                    CLIDRow(j)("Status") = "X"
                                    CLIDRow(j).AcceptChanges()
                                Next
                                myCLIDs.Tables(0).AcceptChanges()
                            End If
                        Next
                    End If
                End If


                'Allow AML Data update in Receiving menu: View Receipt Traveller ?
                GRHeader.AllowAMLUpdate = "NO"
                If GRHeader.Type = 10 Then
                    Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC004'")
                    Dim AMLUpdate As String = Convert.ToString(da.ExecuteScalar(Sqlstr))

                    'If AML Data update is disabled in T_Config, then check user account in table field Exempt
                    If AMLUpdate = "NO" Then
                        Sqlstr = String.Format("Select Exempt from T_Config with (nolock) where ConfigID = 'REC004'")
                        Dim UserExempt As String = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                        If UserExempt <> "" Then
                            If UserExempt.Contains(LoginData.User.ToUpper) Then AMLUpdate = "YES"
                        End If
                    End If
                    GRHeader.AllowAMLUpdate = AMLUpdate
                End If


                Dim HeaderFilled As Boolean = False
                For i = 0 To myCLIDs.Tables(0).Rows.Count - 1
                    If Not myCLIDs.Tables(0).Rows(i)("RecDocItem") Is DBNull.Value Then
                        Dim POLineID As String = myCLIDs.Tables(0).Rows(i)("RecDocItem").ToString          'Identify the unique PO Line   -- 7/3/2017

                        Dim StatusCode As Integer
                        StatusCode = IIf(myCLIDs.Tables(0).Rows(i)("StatusCode") Is DBNull.Value, 0, myCLIDs.Tables(0).Rows(i)("StatusCode"))
                        ProcessID = IIf(myCLIDs.Tables(0).Rows(i)("ProcessID") Is DBNull.Value, 0, myCLIDs.Tables(0).Rows(i)("ProcessID"))

                        If GRHeader.Type = 8 And StatusCode = 0 And myCLIDs.Tables(0).Rows(i)("Status") = "X" Then
                        ElseIf GRHeader.Type = 8 AndAlso ProcessID = 0 Then      'Do not show CLIDs which generated by item transfer ( Set ProcessID = 0 for item transfer )
                        Else
                            Dim POType As String = ""
                            Dim PONo, POItem As String
                            If (ProcessID = 2 Or ProcessID = 9 Or ProcessID = 902) AndAlso myCLIDs.Tables(0).Rows(i)("PurOrdItem") Is DBNull.Value Then
                                POType = "OSP"
                                PONo = myCLIDs.Tables(0).Rows(i)("MatDocNo").ToString
                                POItem = myCLIDs.Tables(0).Rows(i)("MatDocItem").ToString
                            Else
                                PONo = myCLIDs.Tables(0).Rows(i)("PurOrdNo").ToString
                                POItem = myCLIDs.Tables(0).Rows(i)("PurOrdItem").ToString
                            End If

                            If HeaderFilled = False Then
                                If GRHeader.Type <> 9 Then
                                    GRHeader.POCurrency = ""
                                    GRHeader.OrderType = GRHeader.Type
                                    If Not myCLIDs.Tables(0).Rows(i)("InvoiceNo") Is DBNull.Value Then GRHeader.InvoiceNo = myCLIDs.Tables(0).Rows(i)("InvoiceNo")
                                    If Not myCLIDs.Tables(0).Rows(i)("BillOfLading") Is DBNull.Value Then GRHeader.BillOfLading = myCLIDs.Tables(0).Rows(i)("BillOfLading")
                                    If Not myCLIDs.Tables(0).Rows(i)("DN") Is DBNull.Value Then GRHeader.DeliveryNote = myCLIDs.Tables(0).Rows(i)("DN")
                                    If Not myCLIDs.Tables(0).Rows(i)("RecDate") Is DBNull.Value Then GRHeader.PostDate = myCLIDs.Tables(0).Rows(i)("RecDate")
                                    If Not myCLIDs.Tables(0).Rows(i)("HeaderText") Is DBNull.Value Then GRHeader.HeaderText = myCLIDs.Tables(0).Rows(i)("HeaderText")
                                    If Not myCLIDs.Tables(0).Rows(i)("CreatedBy") Is DBNull.Value Then GRHeader.CreatedBy = myCLIDs.Tables(0).Rows(i)("CreatedBy")
                                    If Not myCLIDs.Tables(0).Rows(i)("CreatedOn") Is DBNull.Value Then GRHeader.CreatedOn = myCLIDs.Tables(0).Rows(i)("CreatedOn")
                                    If Not myCLIDs.Tables(0).Rows(i)("ChangedBy") Is DBNull.Value Then GRHeader.ChangedBy = myCLIDs.Tables(0).Rows(i)("ChangedBy")
                                    If Not myCLIDs.Tables(0).Rows(i)("ChangedOn") Is DBNull.Value Then GRHeader.ChangedOn = myCLIDs.Tables(0).Rows(i)("ChangedOn")
                                End If
                                GRHeader.GRNo = GRHeader.OrderNo
                                HeaderFilled = True
                            End If

                            DR = Nothing
                            'DR = POItemData.Select(" PONo = '" & PONo & "' and POItem = '" & POItem & "'")
                            DR = POItemData.Select(" PONo = '" & PONo & "' and POItem = '" & POItem & "' and POLineID = '" & POLineID & "'")
                            If DR.Length = 0 Then
                                PoItemRow = POItemData.NewRow()
                                If Not myCLIDs.Tables(0).Rows(i)("MaterialNo") Is DBNull.Value Then PoItemRow("MaterialNo") = myCLIDs.Tables(0).Rows(i)("MaterialNo")
                                If Not myCLIDs.Tables(0).Rows(i)("MaterialRevision") Is DBNull.Value Then PoItemRow("MaterialRevision") = myCLIDs.Tables(0).Rows(i)("MaterialRevision")
                                If Not myCLIDs.Tables(0).Rows(i)("MaterialDesc") Is DBNull.Value Then PoItemRow("MaterialDesc") = myCLIDs.Tables(0).Rows(i)("MaterialDesc")
                                If Not myCLIDs.Tables(0).Rows(i)("BaseUOM") Is DBNull.Value Then PoItemRow("BaseUOM") = myCLIDs.Tables(0).Rows(i)("BaseUOM")
                                If Not myCLIDs.Tables(0).Rows(i)("VendorPN") Is DBNull.Value Then PoItemRow("VendorPart") = myCLIDs.Tables(0).Rows(i)("VendorPN")
                                If Not myCLIDs.Tables(0).Rows(i)("VendorID") Is DBNull.Value Then PoItemRow("VendorNo") = myCLIDs.Tables(0).Rows(i)("VendorID")
                                If Not myCLIDs.Tables(0).Rows(i)("VendorName") Is DBNull.Value Then PoItemRow("VendorName") = myCLIDs.Tables(0).Rows(i)("VendorName")
                                If Not myCLIDs.Tables(0).Rows(i)("StoragePosition") Is DBNull.Value Then PoItemRow("StoragePosition") = myCLIDs.Tables(0).Rows(i)("StoragePosition")
                                If Not myCLIDs.Tables(0).Rows(i)("ProdDate") Is DBNull.Value Then PoItemRow("ProductionDate") = myCLIDs.Tables(0).Rows(i)("ProdDate")
                                If Not myCLIDs.Tables(0).Rows(i)("DeliveryType") Is DBNull.Value Then PoItemRow("DeliveryType") = myCLIDs.Tables(0).Rows(i)("DeliveryType")
                                If Not myCLIDs.Tables(0).Rows(i)("Operator") Is DBNull.Value Then PoItemRow("Operator") = myCLIDs.Tables(0).Rows(i)("Operator")
                                If Not myCLIDs.Tables(0).Rows(i)("AddlText") Is DBNull.Value Then PoItemRow("ItemText") = myCLIDs.Tables(0).Rows(i)("AddlText")

                                PoItemRow("SLED") = ""
                                If Not myCLIDs.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                                    PoItemRow("ExpDate") = myCLIDs.Tables(0).Rows(i)("ExpDate")
                                    If myCLIDs.Tables(0).Rows(i)("ExpDate").ToString <> "" Then
                                        PoItemRow("SLED") = "X"
                                    End If
                                End If

                                If Not myCLIDs.Tables(0).Rows(i)("MatSuffix1") Is DBNull.Value Then PoItemRow("MatSuffix1") = myCLIDs.Tables(0).Rows(i)("MatSuffix1")
                                If Not myCLIDs.Tables(0).Rows(i)("MatSuffix2") Is DBNull.Value Then PoItemRow("MatSuffix2") = myCLIDs.Tables(0).Rows(i)("MatSuffix2")
                                If Not myCLIDs.Tables(0).Rows(i)("MatSuffix3") Is DBNull.Value Then PoItemRow("MatSuffix3") = myCLIDs.Tables(0).Rows(i)("MatSuffix3")

                                PoItemRow("PONo") = PONo
                                PoItemRow("POItem") = POItem
                                PoItemRow("POLineID") = POLineID             'Identify the unique PO Line   -- 7/3/2017

                                If POType = "OSP" Then
                                    PoItemRow("PullMethod") = "OSP"
                                    PoItemRow("ReasonCode") = myCLIDs.Tables(0).Rows(i)("PurOrdNo").ToString
                                    If GRHeader.Type = 8 Then PoItemRow("ReasonCode") = ""
                                Else
                                    'If myCLIDs.Tables(0).Rows(i)("ProcessID") = 3 Then PoItemRow("PullMethod") = "K"
                                    'If (myCLIDs.Tables(0).Rows(i)("ProcessID") = 301 OrElse myCLIDs.Tables(0).Rows(i)("ProcessID") = 3) Then PoItemRow("PullMethod") = "K"
                                    If ProcessID = 3 OrElse ProcessID = 301 Then
                                        PoItemRow("PullMethod") = "K"
                                    ElseIf ProcessID = 302 Then
                                        PoItemRow("PullMethod") = "EJIT"
                                    ElseIf Left(ProcessID.ToString, 1) = "4" Then
                                        PoItemRow("PullMethod") = "IR"
                                    ElseIf ProcessID = 6 OrElse ProcessID = 7 OrElse ProcessID = 906 OrElse ProcessID = 907 Then
                                        PoItemRow("PullMethod") = "ASN"
                                    End If
                                    If Not myCLIDs.Tables(0).Rows(i)("ReasonCode") Is DBNull.Value Then PoItemRow("ReasonCode") = myCLIDs.Tables(0).Rows(i)("ReasonCode")
                                End If

                                PoItemRow("StatusCode") = StatusCode
                                PoItemRow("Traceable") = IIf(myCLIDs.Tables(0).Rows(i)("IsTraceable") Is DBNull.Value, 0, myCLIDs.Tables(0).Rows(i)("IsTraceable"))
                                PoItemRow("Manufacturer") = myCLIDs.Tables(0).Rows(i)("Manufacturer").ToString
                                PoItemRow("ManufacturerPN") = myCLIDs.Tables(0).Rows(i)("ManufacturerPN").ToString
                                PoItemRow("QMLStatus") = myCLIDs.Tables(0).Rows(i)("QMLStatus").ToString
                                PoItemRow("AddlData") = myCLIDs.Tables(0).Rows(i)("AddlData").ToString
                                PoItemRow("PredefinedSubInv") = myCLIDs.Tables(0).Rows(i)("PredefinedSubInv").ToString
                                PoItemRow("PredefinedLocator") = myCLIDs.Tables(0).Rows(i)("PredefinedLocator").ToString
                                PoItemRow("RTLot") = myCLIDs.Tables(0).Rows(i)("RTLot").ToString

                                PoItemRow("SubInventory") = DBNull.Value
                                PoItemRow("Locator") = DBNull.Value
                                PoItemRow("RevControl") = ""
                                PoItemRow("LotControl") = ""
                                PoItemRow("Stemp") = DBNull.Value
                                PoItemRow("MSL") = DBNull.Value
                                PoItemRow("COC") = DBNull.Value
                                PoItemRow("EER") = DBNull.Value

                                POItemData.Rows.Add(PoItemRow)
                            End If

                            'Fill TraceabilityData
                            If StatusCode = 1 OrElse StatusCode = 9 OrElse (StatusCode = 0 And (GRHeader.Type = 9 OrElse myCLIDs.Tables(0).Rows(i)("Status") = "")) Then
                                TraceabilityDataRow = TraceabilityData.NewRow()
                                TraceabilityDataRow("PONo") = PONo
                                TraceabilityDataRow("POItem") = POItem
                                TraceabilityDataRow("POLineID") = POLineID

                                If Not myCLIDs.Tables(0).Rows(i)("DateCode") Is DBNull.Value Then TraceabilityDataRow("DateCode") = myCLIDs.Tables(0).Rows(i)("DateCode")
                                If Not myCLIDs.Tables(0).Rows(i)("LotNo") Is DBNull.Value Then TraceabilityDataRow("LotNo") = myCLIDs.Tables(0).Rows(i)("LotNo")
                                If Not myCLIDs.Tables(0).Rows(i)("CountryOfOrigin") Is DBNull.Value Then TraceabilityDataRow("COO") = myCLIDs.Tables(0).Rows(i)("CountryOfOrigin")
                                If Not myCLIDs.Tables(0).Rows(i)("QtyBaseUOM") Is DBNull.Value Then TraceabilityDataRow("Qty") = myCLIDs.Tables(0).Rows(i)("QtyBaseUOM")
                                If Not myCLIDs.Tables(0).Rows(i)("CLID") Is DBNull.Value Then TraceabilityDataRow("LID") = myCLIDs.Tables(0).Rows(i)("CLID")
                                If Not myCLIDs.Tables(0).Rows(i)("ReferenceCLID") Is DBNull.Value Then TraceabilityDataRow("RefCLID") = myCLIDs.Tables(0).Rows(i)("ReferenceCLID")
                                If Not myCLIDs.Tables(0).Rows(i)("BoxID") Is DBNull.Value Then TraceabilityDataRow("BoxID") = myCLIDs.Tables(0).Rows(i)("BoxID")
                                If Not myCLIDs.Tables(0).Rows(i)("MaterialNo") Is DBNull.Value Then TraceabilityDataRow("MatNo") = myCLIDs.Tables(0).Rows(i)("MaterialNo")

                                TraceabilityDataRow("SubInventory") = myCLIDs.Tables(0).Rows(i)("SLOC").ToString
                                TraceabilityDataRow("Locator") = myCLIDs.Tables(0).Rows(i)("StorageBin").ToString
                                TraceabilityDataRow("PredefinedSubInv") = myCLIDs.Tables(0).Rows(i)("PredefinedSubInv").ToString
                                TraceabilityDataRow("PredefinedLocator") = myCLIDs.Tables(0).Rows(i)("PredefinedLocator").ToString
                                TraceabilityDataRow("Status") = ""


                                'Get the SubInv / Locator for MergedID, and copy it to the Original CLID    04/23/2013
                                If myCLIDs.Tables(0).Rows(i)("ReferenceCLID") Is DBNull.Value OrElse myCLIDs.Tables(0).Rows(i)("ReferenceCLID").ToString = "" Then
                                ElseIf myCLIDs.Tables(0).Rows(i)("ReferenceCLID").ToString <> "" Then
                                    Dim RefCLID As String = myCLIDs.Tables(0).Rows(i)("ReferenceCLID").ToString
                                    Dim CLIDRow() As DataRow = Nothing
                                    CLIDRow = ds.Tables(0).Select(" CLID = '" & RefCLID & "'")
                                    If CLIDRow.Length > 0 Then
                                        TraceabilityDataRow("SubInventory") = CLIDRow(0)("SLOC").ToString
                                        TraceabilityDataRow("Locator") = CLIDRow(0)("StorageBin").ToString
                                    End If
                                End If


                                'Save TransactionID for later search use only for ReceiptType = 9
                                'If GRHeader.Type = 9 Then
                                '    TraceabilityDataRow("Status") = myCLIDs.Tables(0).Rows(i)("RecDocItem").ToString
                                'End If

                                TraceabilityData.Rows.Add(TraceabilityDataRow)
                            End If
                        End If
                    End If

                Next

                If POItemData.Rows.Count = 0 Then
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow("ErrorMsg") = "^REC-31@"
                    ErrorTable.Rows.Add(ErrorRow)
                End If

            Catch ex As Exception
                ErrorLogging("Receiving-ReadGRItems", LoginData.User.ToUpper, "ReceiptNo: " & GRHeader.OrderNo & " with Receipt type " & GRHeader.Type & ", " & ex.Message & ex.Source, "E")
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "^REC-31@"
                ErrorTable.Rows.Add(ErrorRow)
            End Try
        End Using

    End Function

    Private Function GetItemsFromERP(ByVal LoginData As ERPLogin, ByRef GRHeader As GRHeaderStructure) As DataSet

        GetItemsFromERP = New DataSet
        GetItemsFromERP = AddGRTables()

        Dim POItemData As DataTable
        Dim TraceabilityData As DataTable
        Dim ErrorTable As DataTable
        Dim PoItemRow As Data.DataRow
        Dim TraceabilityDataRow As Data.DataRow
        Dim ErrorRow As Data.DataRow

        'Copy the table structure only
        POItemData = GetItemsFromERP.Tables(0).Clone
        TraceabilityData = GetItemsFromERP.Tables(1).Clone
        ErrorTable = GetItemsFromERP.Tables(2).Clone

        GetItemsFromERP = New DataSet
        GetItemsFromERP.Tables.Add(POItemData)
        GetItemsFromERP.Tables.Add(TraceabilityData)
        GetItemsFromERP.Tables.Add(ErrorTable)

        Dim POTable As DataTable
        Dim PODataRow As Data.DataRow
        Dim myDataColumn As DataColumn

        POTable = New Data.DataTable("POTable")
        myDataColumn = New Data.DataColumn("PONo", System.Type.GetType("System.String"))
        POTable.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("POLine", System.Type.GetType("System.String"))
        POTable.Columns.Add(myDataColumn)

        'p_po_no, p_release_num,p_line_num, p_shipment_num
        Dim POLists As DataSet = Nothing
        If Not GRHeader.POLists Is Nothing Then
            POLists = New DataSet
            POLists = GRHeader.POLists
        End If

        Dim i As Integer
        Dim ErrorMsg As String
        Dim TempStr As String = ""
        Dim PODetails As New DataSet
        Dim myCLIDs As New DataSet
        Dim COOLists As DataSet = Nothing

        Dim DR() As DataRow
        Dim OrderNo, OrderItem As String
        Dim FoundIRData As Boolean = False

        Dim OrderData As POData
        Dim MatlFlag As Boolean = False
        Dim LocatorFlag As Boolean = False

        Dim myLocator As String = ""

        GRHeader.OrderType = GRHeader.Type

        Try

            If Not GRHeader.POLists Is Nothing Then
                Dim Prefix As String = ""
                Dim Sqlstr As String
                Dim tmpCLID As String = ""
                Dim da As DataAccess = GetDataAccess()
                Dim FirstPrefix As String = Left(POLists.Tables(0).Rows(i)("CLID").ToString, 6)
                Dim SecondPrefix As String = Mid(POLists.Tables(0).Rows(i)("CLID").ToString, 3, 4)          '0000

                If GRHeader.CLIDFlag = True Then
                    Sqlstr = String.Format("Select Prefix from T_SEQNo with (nolock) where Type = 'CLID' ")
                    Prefix = Convert.ToString(da.ExecuteScalar(Sqlstr))
                End If

                For i = 0 To POLists.Tables(0).Rows.Count - 1
                    'Check if the upload CLID is valid or not if there has
                    If GRHeader.CLIDFlag = True Then
                        Dim VPCLID As String = POLists.Tables(0).Rows(i)("CLID").ToString
                        If Prefix <> Left(VPCLID, 2) Then
                            ErrorMsg = "Vendor CLID " & VPCLID & " ^REC-109@ " & Prefix                             'VPNCLID Prefix was not started with: LD !
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = ErrorMsg
                            ErrorTable.Rows.Add(ErrorRow)
                            GRHeader.CLIDFlag = False             'Set Flag as False if there has error
                            Exit Function
                        End If

                        If Mid(VPCLID, 3, 1) <> "V" Then
                            ErrorMsg = "Vendor CLID " & VPCLID & ", ^REC-140@"                                       'VPNCLID the third Character must be V.
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = ErrorMsg
                            ErrorTable.Rows.Add(ErrorRow)
                            GRHeader.CLIDFlag = False             'Set Flag as False if there has error
                            Exit Function
                        End If

                        If FirstPrefix <> Left(VPCLID, 6) Then
                            ErrorMsg = "Vendor CLID " & VPCLID & ", ^REC-110@"                                         'VPNCLID Prefix must be the same with the first row of CLID. !
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = ErrorMsg
                            ErrorTable.Rows.Add(ErrorRow)
                            GRHeader.CLIDFlag = False                          'Set Flag as False if there has error
                            Exit Function
                        End If

                        If i = 0 AndAlso SecondPrefix = "0000" Then
                            ErrorMsg = "Vendor CLID " & VPCLID & ", ^REC-111@"                               'VPNCLID Prefix from 3 to 5 must be different with the Prefix from 3 to 5 of local CLID: 000.
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = ErrorMsg
                            ErrorTable.Rows.Add(ErrorRow)
                            GRHeader.CLIDFlag = False                          'Set Flag as False if there has error
                            Exit Function
                        End If


                        Sqlstr = String.Format("Select CLID from T_CLMaster with (nolock) where CLID = '{0}'", VPCLID)
                        tmpCLID = Convert.ToString(da.ExecuteScalar(Sqlstr))
                        If tmpCLID <> "" Then
                            ErrorMsg = "Vendor CLID " & VPCLID & ", ^REC-112@"                                 'VPNCLID already exists in eTrace !
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = ErrorMsg
                            ErrorTable.Rows.Add(ErrorRow)
                            GRHeader.CLIDFlag = False             'Set Flag as False if there has error
                            Exit Function
                        End If
                    End If

                    If POLists.Tables(0).Rows(i)(0) Is DBNull.Value Then
                    ElseIf POLists.Tables(0).Rows(i)(0) = "" Then
                    Else
                        OrderNo = POLists.Tables(0).Rows(i)("PONo")
                        OrderItem = POLists.Tables(0).Rows(i)("POLine")

                        DR = Nothing
                        DR = POTable.Select(" PONo = '" & OrderNo & "' and POLine = '" & OrderItem & "'")
                        If DR.Length = 0 Then
                            PODataRow = POTable.NewRow()
                            PODataRow("PONo") = OrderNo
                            PODataRow("POLine") = OrderItem
                            POTable.Rows.Add(PODataRow)
                        End If
                    End If
                Next

                If GRHeader.Type = 1 Or GRHeader.Type = 2 Then
                    PODetails = Get_Batch_PO(LoginData.OrgCode, POTable)
                ElseIf GRHeader.Type = 3 Then
                    PODetails = Get_Batch_EJITPO(LoginData.OrgCode, POTable)
                ElseIf GRHeader.Type = 6 Or GRHeader.Type = 7 Then
                    PODetails = Get_Batch_ASN(LoginData.OrgCode, POTable)
                End If

            ElseIf GRHeader.Type = 9 Then             'Read RT Data base on Receipt Number
                PODetails = Get_RT_Data(LoginData.OrgCode, GRHeader.OrderNo)
            Else

                If GRHeader.Type = 6 Or GRHeader.Type = 7 Then
                    OrderData = Split_POData("", GRHeader.OrderItem)
                    OrderData.PONo = GRHeader.OrderNo
                Else
                    OrderData = Split_POData(GRHeader.OrderNo, GRHeader.OrderItem)
                End If
                If OrderData.MaterialNo <> "" Then MatlFlag = True

                If OrderData.PONo = "" Then
                    ErrorMsg = "Invalid PO " & GRHeader.OrderNo
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Read Locator first from ProductionLine, then from column POLine (User input Locator start with Lxxxxxxx )
                If GRHeader.Type = 3 Or GRHeader.Type = 4 Or GRHeader.Type = 7 Then
                    If LoginData.ProductionLine.ToString <> "" Then
                        LocatorFlag = True
                        myLocator = LoginData.ProductionLine.ToString.ToUpper
                    Else
                        If OrderData.Locator <> "" Then
                            LocatorFlag = True
                            myLocator = OrderData.Locator.ToUpper
                        End If
                    End If
                End If


                If GRHeader.Type = 3 Then
                    PODetails = Get_PO_Kanban(LoginData.OrgCode, OrderData)
                ElseIf GRHeader.Type = 4 Then
                    PODetails = Get_IR_ISO(LoginData.OrgCode, OrderData)
                ElseIf GRHeader.Type = 5 Then
                    PODetails = Get_PO_Indirect(LoginData.OrgCode, OrderData)
                ElseIf GRHeader.Type = 6 Or GRHeader.Type = 7 Then
                    PODetails = Get_ASN_PO(LoginData.OrgCode, OrderData)
                Else
                    PODetails = Get_PO_Detail(LoginData.OrgCode, GRHeader.Type, OrderData)
                End If
            End If

            If PODetails Is Nothing OrElse PODetails.Tables.Count = 0 OrElse PODetails.Tables("po_header").Rows.Count = 0 _
            OrElse PODetails.Tables("po_line").Rows.Count = 0 OrElse PODetails.Tables("po_header").Columns.Count = 1 _
            OrElse PODetails.Tables("po_line").Columns.Count = 1 Then
                If (GRHeader.Type = 1 Or GRHeader.Type = 3) And Not POLists Is Nothing Then
                    ErrorMsg = "Invalid PO Data from your upload Excel file"
                ElseIf (GRHeader.Type = 6 Or GRHeader.Type = 7) And Not POLists Is Nothing Then
                    ErrorMsg = "Invalid ASN Data from your upload file: ASN No# and ASN Line No# are not matched."
                ElseIf GRHeader.Type = 2 And Not POLists Is Nothing Then
                    ErrorMsg = "Invalid OSP PO Data from your upload Excel file"
                ElseIf GRHeader.Type = 9 Then
                    ErrorMsg = "Invalid Receipt Number " & GRHeader.OrderNo
                ElseIf GRHeader.Type = 3 Then
                    ErrorMsg = "Non-KanBan PO or Invalid PO " & GRHeader.OrderNo
                ElseIf GRHeader.Type = 4 Then
                    ErrorMsg = "Invalid IR/ISO ShipmentNo " & GRHeader.OrderNo
                ElseIf GRHeader.Type = 5 Then
                    ErrorMsg = "Invalid Indirect PO " & GRHeader.OrderNo
                ElseIf GRHeader.Type = 6 OrElse GRHeader.Type = 7 Then
                    ErrorMsg = "Invalid ASN No# " & GRHeader.OrderNo
                Else
                    ErrorMsg = "Invalid PO " & GRHeader.OrderNo
                End If

                If GRHeader.OrderItem Is Nothing Then
                ElseIf GRHeader.OrderItem.ToString <> "" And (Not GRHeader.OrderItem.ToString.StartsWith("P")) Then
                    ErrorMsg = ErrorMsg & " / " & GRHeader.OrderItem
                End If

                ErrorRow = ErrorTable.NewRow()
                ErrorRow(0) = ErrorMsg
                ErrorTable.Rows.Add(ErrorRow)
                Exit Function
            End If

            If GRHeader.Type < 8 Then
                If Not POLists Is Nothing AndAlso PODetails.Tables("po_header").Rows.Count > 1 Then
                    Dim VendorNo As String = PODetails.Tables("po_header").Rows(0)("VENDOR_NUMBER")
                    DR = Nothing
                    DR = PODetails.Tables("po_header").Select(" VENDOR_NUMBER <> '" & VendorNo & "'")
                    If DR.Length > 0 Then
                        ErrorMsg = "^REC-29@"                                 'Multiple vendors not allowed !
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = ErrorMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        Exit Function
                    End If
                End If
            End If

            'Check if there has OSP DJs which have OQA operation before DJ completion or not
            If GRHeader.Type = 2 AndAlso PODetails.Tables("po_line").Rows.Count > 1 Then
                Dim OQAFlag As String = PODetails.Tables("po_line").Rows(0)("oqa_flag").ToString
                DR = Nothing
                DR = PODetails.Tables("po_line").Select(" oqa_flag <> '" & OQAFlag & "'")
                If DR.Length > 0 Then
                    ErrorMsg = "^REC-88@"  '"Please check your input data and make sure all OSP DJs have OQA operation or none DJ have"                                 'Multiple vendors not allowed !
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If
            End If


            'Filter data according to user's input Material Number 
            If MatlFlag = True Then
                DR = Nothing
                If GRHeader.Type = 2 Then
                    DR = PODetails.Tables("po_line").Select(" WIP_ASSEMBLY <> '" & OrderData.MaterialNo & "'")
                Else
                    DR = PODetails.Tables("po_line").Select(" ITEM_NUMBER <> '" & OrderData.MaterialNo & "'")
                End If
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    PODetails.Tables("po_line").AcceptChanges()

                    If PODetails.Tables("po_line").Rows.Count = 0 Then
                        If GRHeader.Type = 3 Then
                            ErrorMsg = "Couldn't find KB/eJit PO " & GRHeader.OrderNo & " with the item: " & OrderData.MaterialNo
                        ElseIf GRHeader.Type = 4 Then
                            ErrorMsg = "Couldn't find IR/ISO ShipmentNo " & GRHeader.OrderNo & " with the item: " & OrderData.MaterialNo
                        ElseIf GRHeader.Type = 6 OrElse GRHeader.Type = 7 Then
                            ErrorMsg = "Couldn't find ASN " & GRHeader.OrderNo & " with the item: " & OrderData.MaterialNo
                        Else
                            ErrorMsg = "Couldn't find PO " & GRHeader.OrderNo & " with the item: " & OrderData.MaterialNo
                        End If
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = ErrorMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        Exit Function
                    End If
                End If
            End If


            'Filter data according to user's input Locator 
            If LocatorFlag = True Then
                DR = Nothing
                DR = PODetails.Tables("po_line").Select(" DESTINATION_LOCATOR <> '" & myLocator & "'")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        DR(i).Delete()
                    Next
                    PODetails.Tables("po_line").AcceptChanges()

                    If PODetails.Tables("po_line").Rows.Count = 0 Then
                        If GRHeader.Type = 3 Then
                            ErrorMsg = "Couldn't find KB/eJit PO " & GRHeader.OrderNo & " with your Locator: " & myLocator
                        ElseIf GRHeader.Type = 7 Then
                            ErrorMsg = "Couldn't find ASN " & GRHeader.OrderNo & " with your Locator: " & myLocator
                        Else
                            ErrorMsg = "Couldn't find IR/ISO ShipmentNo " & GRHeader.OrderNo & " with your Locator: " & myLocator
                        End If
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = ErrorMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        Exit Function
                    End If
                End If
            End If


            'Identify Order Type according to RT Data
            If GRHeader.Type = 9 Then
                If PODetails.Tables("po_header").Rows(0)("receipt_type").ToString.ToUpper = "IR" Then
                    GRHeader.OrderType = 4
                Else
                    If PODetails.Tables("po_line").Rows(0)("line_type_id") = 3 Then            'OSP 
                        GRHeader.OrderType = 2
                    Else                                                                       'Goods
                        If Mid(PODetails.Tables("po_line").Rows(0)("PO_NUMBER"), 8, 1) = "4" OrElse _
                           Mid(PODetails.Tables("po_line").Rows(0)("PO_NUMBER"), 8, 1) = "5" Then
                            GRHeader.OrderType = 3
                            If FixNull(PODetails.Tables("po_line").Rows(0)("asn_type")) <> "" Then              'RT created by ASN Receipt
                                GRHeader.OrderType = 7
                            End If
                        Else
                            GRHeader.OrderType = 1
                            If FixNull(PODetails.Tables("po_line").Rows(0)("asn_type")) <> "" Then              'RT created by ASN Receipt
                                GRHeader.OrderType = 6
                            End If
                        End If
                    End If
                End If
            End If

            'Read MSL information from iPro for Normal PO    01/10/2015
            Dim dsAML As DataSet = New DataSet
            If GRHeader.Type = 1 Or GRHeader.Type = 6 Then       'Normal PO / ASN              
                Dim j, k As Integer
                Dim MatList() As String
                j = PODetails.Tables("po_line").Rows.Count
                ReDim MatList(j)
                MatList.Initialize()

                k = 0
                For i = 0 To PODetails.Tables("po_line").Rows.Count - 1
                    MatList(k) = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                    k = k + 1
                Next

                Try
                    dsAML = GetAMLData(MatList, "RECEIVING")
                Catch ex As Exception
                    ErrorLogging("Receiving-GetAMLData", LoginData.User.ToUpper, "OrderNo: " & GRHeader.OrderNo & " with Receipt type " & GRHeader.Type & ", " & ex.Message & ex.Source, "E")
                End Try
            End If


            If GRHeader.OrderType = 4 Then
                Dim myIRData As IRData = New IRData
                myIRData.User = LoginData.User.ToUpper
                myIRData.SourceOrg = PODetails.Tables("po_header").Rows(0)("VENDOR_NUMBER").ToString
                myIRData.DestOrg = LoginData.OrgCode

                If GRHeader.Type = 4 Then
                    myIRData.ShipmentNo = GRHeader.OrderNo              'DN Number
                Else
                    myIRData.ShipmentNo = PODetails.Tables("po_line").Rows(0)("SHIPMENT_NUM").ToString
                End If
                myCLIDs = GetIRDataFromServer(myIRData)

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                ElseIf myCLIDs.Tables(0).Rows.Count > 0 Then
                    'Get CLIDs which already done receipt in DestOrg Server
                    Dim ShipCLIDs As DataSet = New DataSet
                    ShipCLIDs = ReadShipCLIDs(LoginData, GRHeader.OrderNo)

                    If ShipCLIDs Is Nothing OrElse ShipCLIDs.Tables.Count = 0 Then
                    ElseIf ShipCLIDs.Tables(0).Rows.Count > 0 Then
                        For i = 0 To ShipCLIDs.Tables(0).Rows.Count - 1
                            Dim ShipCLID As String = ShipCLIDs.Tables(0).Rows(i)("CLID").ToString
                            DR = Nothing
                            DR = myCLIDs.Tables(0).Select(" CLID = '" & ShipCLID & "'")
                            If DR.Length > 0 Then
                                DR(0).Delete()
                            End If
                        Next
                        myCLIDs.Tables(0).AcceptChanges()
                    End If
                    If myCLIDs.Tables(0).Rows.Count > 0 Then FoundIRData = True 'CLID found from Source Server

                End If
            End If

            GRHeader.PostDate = Date.Today
            GRHeader.POCurrency = ""
            GRHeader.POCurrency = PODetails.Tables("po_header").Rows(0)("currency_code").ToString
            'If GRHeader.Type = 1 OrElse GRHeader.Type = 6 OrElse GRHeader.Type = 9 Then
            '    GRHeader.POCurrency = PODetails.Tables("po_header").Rows(0)("currency_code").ToString
            'End If

            'Set default value for DeliveryType if the Currency is CNY with FY & LD    09/27/2013
            Dim DeliveryType As String = ""
            If (GRHeader.OrderType = 1 Or GRHeader.OrderType = 3 Or GRHeader.OrderType = 6 Or GRHeader.OrderType = 7) _
                 AndAlso GRHeader.POCurrency = "CNY" Then
                Dim Sqlstr As String
                Dim da As DataAccess = GetDataAccess()
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC005'")
                DeliveryType = Convert.ToString(da.ExecuteScalar(Sqlstr))
            End If

            If GRHeader.Type = 6 OrElse GRHeader.Type = 7 OrElse GRHeader.Type = 9 Then
                If GRHeader.Type = 9 Then
                    GRHeader.GRNo = GRHeader.OrderNo
                    GRHeader.PostDate = CDate(PODetails.Tables("po_header").Rows(0)("receipt_date"))
                End If
                GRHeader.InvoiceNo = PODetails.Tables("po_header").Rows(0)("packing_slip").ToString
                GRHeader.BillOfLading = PODetails.Tables("po_header").Rows(0)("bill_of_lading").ToString
                GRHeader.DeliveryNote = PODetails.Tables("po_header").Rows(0)("waybill_airbill_num").ToString
                GRHeader.HeaderText = PODetails.Tables("po_header").Rows(0)("comments").ToString
            End If

            'Add new column POLineID to iendtify the record in table po_line for later search use     8/24/2011
            PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("POLineID", System.Type.GetType("System.String")))

            For i = 0 To PODetails.Tables("po_line").Rows.Count - 1
                If PODetails.Tables("po_line").Rows(i)("RELEASE_NUM") Is DBNull.Value Then PODetails.Tables("po_line").Rows(i)("RELEASE_NUM") = 0
                If PODetails.Tables("po_line").Rows(i)("SHIPMENT_NUM") Is DBNull.Value Then PODetails.Tables("po_line").Rows(i)("SHIPMENT_NUM") = 0

                Dim myPN As String = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                OrderNo = PODetails.Tables("po_line").Rows(i)("PO_NUMBER").ToString
                OrderItem = PODetails.Tables("po_line").Rows(i)("LINE_NUM").ToString

                If PODetails.Tables("po_line").Rows(i)("RELEASE_NUM") > 0 Then
                    OrderNo = PODetails.Tables("po_line").Rows(i)("PO_NUMBER").ToString & "-" & PODetails.Tables("po_line").Rows(i)("RELEASE_NUM").ToString
                End If
                If GRHeader.OrderType = 4 Then OrderNo = PODetails.Tables("po_line").Rows(i)("SHIPMENT_NUM").ToString
                If PODetails.Tables("po_line").Rows(i)("SHIPMENT_NUM") > 0 AndAlso GRHeader.OrderType <> 4 Then
                    OrderItem = PODetails.Tables("po_line").Rows(i)("LINE_NUM").ToString & "." & PODetails.Tables("po_line").Rows(i)("SHIPMENT_NUM").ToString
                End If


                'Check if users choose wrong receipt type for normal PO
                If (GRHeader.Type <> 2 And GRHeader.Type <> 9) AndAlso PODetails.Tables("po_line").Rows(i)("line_type_id") = 3 Then     'Outside Processing
                    ErrorMsg = "You should choose Receipt Type 2 for OSP PO " & OrderNo & "/" & OrderItem
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if users choose wrong receipt type for OSP PO
                If GRHeader.Type = 2 AndAlso PODetails.Tables("po_line").Rows(i)("line_type_id") = 1 Then     'Goods
                    TempStr = myPN & " with PO " & OrderNo & "/" & OrderItem
                    ErrorMsg = "You should choose Receipt Type 1 for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                If PODetails.Tables("po_line").Rows(i)("receiving_routing_id") Is DBNull.Value OrElse PODetails.Tables("po_line").Rows(i)("receiving_routing_id").ToString = "" Then
                    TempStr = myPN & " with PO " & OrderNo & "/" & OrderItem
                    ErrorMsg = "Blank Receipt Routing is not allowed for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if this ASN is for Normal PO, or EJIT PO, user should choose the correct Receipt Type
                If GRHeader.Type = 6 AndAlso PODetails.Tables("po_line").Rows(i)("po_type") <> "NORMAL_ASNPO" Then
                    Dim ASNNo, ASNLine As String
                    ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                    ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString

                    TempStr = myPN & " with ASN " & ASNNo & "/" & ASNLine
                    ErrorMsg = "This ASN is for KB/eJit PO. You should choose Receipt Type 7 for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if this ASN is for Normal PO, or EJIT PO, user should choose the correct Receipt Type
                If GRHeader.Type = 7 AndAlso PODetails.Tables("po_line").Rows(i)("po_type") = "NORMAL_ASNPO" Then
                    Dim ASNNo, ASNLine As String
                    ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                    ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString

                    TempStr = myPN & " with ASN " & ASNNo & "/" & ASNLine
                    ErrorMsg = "This ASN is for Normal PO. You should choose Receipt Type 6 for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if users choose wrong receipt type for Normal PO
                If (GRHeader.Type = 1 OrElse GRHeader.Type = 6) AndAlso PODetails.Tables("po_line").Rows(i)("receiving_routing_id") = 3 Then    'Direct Delivery
                    TempStr = myPN & " with Normal PO " & OrderNo & "/" & OrderItem
                    ErrorMsg = "Wrong Receipt Routing: Direct Delivery for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Not allow to receive PO Item which already exists ASN for this PO shipment
                If (GRHeader.Type = 1 Or GRHeader.Type = 3) AndAlso FixNull(PODetails.Tables("po_line").Rows(i)("asn_type")) <> "" Then
                    If Mid(OrderNo, 8, 1) = "4" OrElse Mid(OrderNo, 8, 1) = "5" Then       'KB/eJit PO
                        TempStr = myPN & " with KB/eJit PO " & OrderNo & "/" & OrderItem
                        ErrorMsg = "An ASN already exists for this PO shipment. You should choose Receipt Type 7 for Item " & TempStr
                    Else
                        TempStr = myPN & " with Normal PO " & OrderNo & "/" & OrderItem
                        ErrorMsg = "An ASN already exists for this PO shipment. You should choose Receipt Type 6 for Item " & TempStr
                    End If
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if the KB/eJit PO has wrong receipt routing, it should be Direct Delivery
                If (GRHeader.Type = 3 Or GRHeader.Type = 7) AndAlso PODetails.Tables("po_line").Rows(i)("receiving_routing_id") <> 3 Then
                    TempStr = myPN & " with KB/eJit PO " & OrderNo & "/" & OrderItem
                    If GRHeader.Type = 7 Then
                        Dim ASNNo, ASNLine As String
                        ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                        ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString
                        TempStr = TempStr & ",  the ASN Line is " & ASNNo & "/" & ASNLine
                    End If
                    ErrorMsg = "Wrong Receipt Routing: " & PODetails.Tables("po_line").Rows(i)("receiving_routing") & " for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if the Kanban PO has its default SubInv / Locator
                If (GRHeader.Type = 3 Or GRHeader.Type = 7) AndAlso PODetails.Tables("po_line").Rows(i)("receiving_routing_id") = 3 Then
                    'Fix EJIT/Kanban ASN Destination Locator issue 2018-06-11
                    If GRHeader.Type = 7 And PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString = "" And PODetails.Tables("po_line").Rows(i)("LOCATOR_ID").ToString <> "" Then
                        Dim Sqlstr As String
                        Dim DestinationLocator As String = ""
                        Dim da As DataAccess = GetDataAccess()
                        DR = Nothing
                        DR = PODetails.Tables("po_line").Select(" LOCATOR_ID = " & PODetails.Tables("po_line").Rows(i)("LOCATOR_ID").ToString & " and DESTINATION_LOCATOR <> ''")
                        If DR.Length = 0 Then
                            Sqlstr = String.Format("exec dbo.ora_EJITASN_Locator '{0}','{1}'", PODetails.Tables("po_line").Rows(i)("LOCATOR_ID").ToString, LoginData.OrgID.ToString)
                            DestinationLocator = Convert.ToString(da.ExecuteScalar(Sqlstr))
                        Else
                            DestinationLocator = DR(0)("DESTINATION_LOCATOR")
                        End If
                        If DestinationLocator <> "" Then
                            PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR") = DestinationLocator
                        End If
                    End If
                    If PODetails.Tables("po_line").Rows(i)("DESTINATION_SUBINVENTORY").ToString = "" OrElse _
                        PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString = "" Then
                        TempStr = myPN & " with KB/eJit PO " & OrderNo & "/" & OrderItem
                        If GRHeader.Type = 7 Then
                            Dim ASNNo, ASNLine As String
                            ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                            ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString
                            TempStr = TempStr & ",  the ASN Line is " & ASNNo & "/" & ASNLine
                        End If
                        ErrorMsg = "Can't find its default SubInv or Locator for Item " & TempStr
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = ErrorMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        Exit Function
                    End If
                End If


                'Check if the Normal IR has wrong receipt routing: Direct Delivery, it should be Standard Receipt / Inspection Required
                If GRHeader.Type = 4 AndAlso PODetails.Tables("po_line").Rows(i)("ir_type") = "NO_IR" AndAlso _
                   PODetails.Tables("po_line").Rows(i)("receiving_routing_id") = 3 Then
                    TempStr = myPN & " with ShipmentNo " & OrderNo & "/" & OrderItem
                    ErrorMsg = "Wrong Receipt Routing: Direct Delivery for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If

                'Check if the KB_IR / EJ_IR has wrong receipt routing, it should be Direct Delivery
                If GRHeader.Type = 4 AndAlso PODetails.Tables("po_line").Rows(i)("ir_type") <> "NO_IR" AndAlso _
                   PODetails.Tables("po_line").Rows(i)("receiving_routing_id") <> 3 Then
                    TempStr = myPN & " with ShipmentNo " & OrderNo & "/" & OrderItem
                    ErrorMsg = "Wrong Receipt Routing: " & PODetails.Tables("po_line").Rows(i)("receiving_routing") & " for Item " & TempStr
                    ErrorRow = ErrorTable.NewRow()
                    ErrorRow(0) = ErrorMsg
                    ErrorTable.Rows.Add(ErrorRow)
                    Exit Function
                End If


                Dim eJitPOLine As String = OrderItem      'Record eJitPOLine and Shipment for later search
                If GRHeader.OrderType = 3 Or GRHeader.OrderType = 7 Then OrderItem = OrderItem & "." & PODetails.Tables("po_line").Rows(i)("DISTRIBUTION_NUM").ToString

                Dim POLineID As String = ""
                If GRHeader.Type = 1 OrElse GRHeader.Type = 2 OrElse GRHeader.Type = 6 Then
                    POLineID = PODetails.Tables("po_line").Rows(i)("PO_LINE_LOCATION_ID").ToString
                ElseIf GRHeader.Type = 3 OrElse GRHeader.Type = 5 OrElse GRHeader.Type = 7 Then
                    POLineID = PODetails.Tables("po_line").Rows(i)("PO_DISTRIBUTION_ID").ToString
                ElseIf GRHeader.Type = 4 Then
                    POLineID = PODetails.Tables("po_line").Rows(i)("SHIPMENT_LINE_ID").ToString
                ElseIf GRHeader.Type = 9 Then
                    POLineID = PODetails.Tables("po_line").Rows(i)("TRANSACTION_ID").ToString
                End If
                'Save the value in the new column POLineID in table po_line for later search use     8/24/2011
                PODetails.Tables("po_line").Rows(i)("POLineID") = POLineID
                PODetails.Tables("po_line").Rows(i).AcceptChanges()

                'Check if there has duplicate PO Line data  01/04/2012
                Dim R1() As DataRow = Nothing
                If GRHeader.Type = 9 Then
                    R1 = POItemData.Select(" PONo = '" & OrderNo & "' and POItem = '" & OrderItem & "' and POLineID = '" & POLineID & "'")
                Else
                    R1 = POItemData.Select(" PONo = '" & OrderNo & "' and POItem = '" & OrderItem & "'")
                End If
                If R1.Length > 0 Then Continue For

                'Fill in PO data
                PoItemRow = POItemData.NewRow()
                PoItemRow("PONo") = OrderNo
                PoItemRow("POItem") = OrderItem
                PoItemRow("POLineID") = POLineID
                PoItemRow("UOM") = PODetails.Tables("po_line").Rows(i)("PO_UOM").ToString
                PoItemRow("BaseUOM") = PODetails.Tables("po_line").Rows(i)("BASE_UOM_CODE").ToString
                PoItemRow("VendorNo") = PODetails.Tables("po_line").Rows(i)("VENDOR_NUMBER").ToString
                PoItemRow("VendorName") = PODetails.Tables("po_line").Rows(i)("VENDOR").ToString
                PoItemRow("EarlyDays") = PODetails.Tables("po_line").Rows(i)("DAYS_EARLY_RECEIPT_ALLOWED")
                PoItemRow("VendorPart") = ""
                PoItemRow("SLED") = ""
                PoItemRow("ExpDays") = 0
                PoItemRow("MinRemDays") = 0
                PoItemRow("BaseUOMConv") = 1    'PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE")
                PoItemRow("PullMethod") = ""
                PoItemRow("MaterialRevision") = ""
                PoItemRow("NeedbyDate") = DBNull.Value
                PoItemRow("PromisedDate") = DBNull.Value
                PoItemRow("PredefinedSubInv") = DBNull.Value
                PoItemRow("PredefinedLocator") = DBNull.Value
                PoItemRow("StockType") = "FTS"
                PoItemRow("RTLot") = ""

                PoItemRow("RoHS") = ""
                PoItemRow("Stemp") = ""
                PoItemRow("MSL") = ""
                PoItemRow("COC") = ""
                PoItemRow("EER") = ""
                PoItemRow("AddlData") = ""

                'Get Lot Control info / Revision Control 
                PoItemRow("RevControl") = "N"
                PoItemRow("LotControl") = "N"
                PoItemRow("ExpControl") = "N"
                If Not PODetails.Tables("po_line").Rows(i)("revision_control_code") Is DBNull.Value Then
                    If PODetails.Tables("po_line").Rows(i)("revision_control_code") = 2 Then PoItemRow("RevControl") = "Y"
                End If

                'Only the Lot control is turned on, then we will look at Expiration Date Control, otherwise, set ExpControl = "N"
                If Not PODetails.Tables("po_line").Rows(i)("lot_control_code") Is DBNull.Value Then
                    If PODetails.Tables("po_line").Rows(i)("lot_control_code") = 2 Then
                        PoItemRow("LotControl") = "Y"
                        'shelf_life_code only used for Non OSP PO type   08/20/2012 
                        If GRHeader.OrderType <> 2 AndAlso (Not PODetails.Tables("po_line").Rows(i)("shelf_life_code") Is DBNull.Value) Then
                            If PODetails.Tables("po_line").Rows(i)("shelf_life_code") <> 1 Then PoItemRow("ExpControl") = "Y"
                        End If
                    End If
                End If

                Dim MDate As Date = New Date
                If Not PODetails.Tables("po_line").Rows(i)("NEED_BY_DATE") Is DBNull.Value Then
                    MDate = CDate(PODetails.Tables("po_line").Rows(i)("NEED_BY_DATE"))
                    PoItemRow("NeedbyDate") = MDate
                End If

                'Add PO Promised Date    -   08/19/2013
                If GRHeader.OrderType <> 4 Then
                    'If GRHeader.OrderType <> 4 AndAlso GRHeader.Type <> 2 Then
                    If Not PODetails.Tables("po_line").Rows(i)("PO_PROMISED_DATE") Is DBNull.Value Then
                        Dim PDate As Date = New Date
                        PDate = CDate(PODetails.Tables("po_line").Rows(i)("PO_PROMISED_DATE"))
                        PoItemRow("PromisedDate") = PDate
                    End If
                End If

                'Set default value for DeliveryType if the Currency is CNY with FY & LD    09/27/2013
                If (GRHeader.OrderType = 1 Or GRHeader.OrderType = 3 Or GRHeader.OrderType = 6 Or GRHeader.OrderType = 7) _
                   AndAlso GRHeader.POCurrency = "CNY" Then
                    If DeliveryType <> "" Then PoItemRow("DeliveryType") = DeliveryType '   VAT
                End If

                If GRHeader.OrderType = 2 Then
                    PoItemRow("MaterialNo") = PODetails.Tables("po_line").Rows(i)("WIP_ASSEMBLY").ToString
                    PoItemRow("MaterialDesc") = PODetails.Tables("po_line").Rows(i)("WIP_DESCRIPTION").ToString
                    PoItemRow("ReasonCode") = PODetails.Tables("po_line").Rows(i)("WIP_ENTITY_NAME").ToString                  'OSP DJ
                    PoItemRow("SubInventory") = PODetails.Tables("po_line").Rows(i)("WIP_COMPLETION_SUBINVENTORY").ToString    'OSP DJ SubInventory
                    PoItemRow("Locator") = PODetails.Tables("po_line").Rows(i)("WIP_COMPLETION_LOCATOR").ToString              'OSP DJ Locator
                    PoItemRow("RoHS") = PODetails.Tables("po_line").Rows(i)("ROHS").ToString                                   'ROHS Flag
                    PoItemRow("PullMethod") = "OSP"                 'OSP DJ

                    'Read OSP Assembly UOM from WIP UOM        02/22/2016
                    PoItemRow("UOM") = PODetails.Tables("po_line").Rows(i)("WIP_UOM").ToString
                    PoItemRow("BaseUOM") = PODetails.Tables("po_line").Rows(i)("WIP_UOM_CODE").ToString

                    If GRHeader.Type = 9 Then
                        PoItemRow("QtyBaseUOM") = PODetails.Tables("po_line").Rows(i)("RT_QTY") * PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE")
                    Else
                        PoItemRow("QtyBaseUOM") = PODetails.Tables("po_line").Rows(i)("BALANCE_QTY")
                    End If
                    PoItemRow("GRQty") = PoItemRow("QtyBaseUOM")

                    If PoItemRow("RevControl") = "Y" Then
                        If PODetails.Tables("po_line").Rows(i)("WIP_REVISION").ToString <> "" Then
                            PoItemRow("MaterialRevision") = PODetails.Tables("po_line").Rows(i)("WIP_REVISION").ToString
                        Else
                            PoItemRow("MaterialRevision") = PODetails.Tables("po_line").Rows(i)("item_rev").ToString
                        End If
                    End If

                Else
                    'Set Default CONVERSION_RATE as 1 if no value found
                    If PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE") Is DBNull.Value Then PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE") = 1

                    PoItemRow("MaterialNo") = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                    PoItemRow("MaterialDesc") = PODetails.Tables("po_line").Rows(i)("ITEM_DESCRIPTION").ToString
                    PoItemRow("VendorPart") = PODetails.Tables("po_line").Rows(i)("VENDOR_PRODUCT_NUM").ToString
                    PoItemRow("BaseUOMConv") = PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE") 'conversion_rate
                    PoItemRow("SubInventory") = DBNull.Value
                    PoItemRow("Locator") = DBNull.Value
                    PoItemRow("ReasonCode") = ""

                    'Read RoHS / Stmp / MSL / COC / ESD / EER from iPro for Normal PO   -- 10/12/2016
                    If dsAML Is Nothing OrElse dsAML.Tables.Count = 0 Then
                    ElseIf dsAML.Tables("ItemData").Rows.Count > 0 Then
                        DR = Nothing
                        Dim MaterialNo As String = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                        DR = dsAML.Tables("ItemData").Select(" MaterialNo = '" & MaterialNo & "'")
                        If DR.Length > 0 Then
                            PoItemRow("RoHS") = IIf(DR(0)("RoHS") Is DBNull.Value, "", DR(0)("RoHS"))
                            PoItemRow("Stemp") = IIf(DR(0)("Stemp") Is DBNull.Value, "", DR(0)("Stemp"))
                            PoItemRow("MSL") = IIf(DR(0)("MSL") Is DBNull.Value, "", DR(0)("MSL"))
                            PoItemRow("COC") = IIf(DR(0)("COC") Is DBNull.Value, "", DR(0)("COC"))
                            'PoItemRow("EER") = IIf(DR(0)("EER") Is DBNull.Value, "", DR(0)("EER"))                    'Add EER from iPro here
                            PoItemRow("AddlData") = IIf(DR(0)("AddlData") Is DBNull.Value, "", DR(0)("AddlData"))
                            PoItemRow("ReasonCode") = PoItemRow("MSL")                'Default MSL in ReasonCode column for normal PO
                            PoItemRow("RESData") = DR(0)("RESData")
                        End If
                    End If


                    If GRHeader.OrderType = 3 OrElse GRHeader.OrderType = 4 OrElse GRHeader.OrderType = 7 Then
                        PoItemRow("SubInventory") = PODetails.Tables("po_line").Rows(i)("DESTINATION_SUBINVENTORY").ToString
                        PoItemRow("Locator") = PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString
                        If GRHeader.Type <> 9 Then PoItemRow("ReasonCode") = PODetails.Tables("po_line").Rows(i)("KANBAN_CARD_ID").ToString

                        If GRHeader.OrderType = 3 Or GRHeader.OrderType = 7 Then PoItemRow("PullMethod") = "K"
                        If GRHeader.OrderType = 4 Then
                            PoItemRow("PullMethod") = "IR"
                            PoItemRow("RTLot") = PODetails.Tables("po_line").Rows(i)("RT_LOT_NUMBER").ToString
                        End If
                    End If

                    'Header.Type = 9:  CLID creation based on Oracle Receipt
                    If GRHeader.Type = 9 Then
                        PoItemRow("QtyBaseUOM") = PODetails.Tables("po_line").Rows(i)("RT_QTY") * PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE")
                        PoItemRow("GRQty") = PoItemRow("QtyBaseUOM")
                        If DeliveryType = "" AndAlso PoItemRow("DeliveryType") Is DBNull.Value Then
                            PoItemRow("DeliveryType") = PODetails.Tables("po_line").Rows(i)("DELIVERY_TYPE").ToString
                        End If

                        Dim Comments As String = PODetails.Tables("po_line").Rows(i)("COMMENTS").ToString.Trim
                        If Comments <> "" Then
                            Dim Arry() As String = Split(Comments, " ")
                            If Arry.Length = 1 Then
                                PoItemRow("Operator") = Arry(0).ToString
                            ElseIf Arry.Length = 2 Then
                                PoItemRow("Operator") = Arry(0).ToString
                                PoItemRow("StoragePosition") = Arry(1).ToString
                            ElseIf Arry.Length = 3 Then
                                PoItemRow("Operator") = Arry(0).ToString
                                PoItemRow("StoragePosition") = Arry(1).ToString
                                PoItemRow("BoxDesc") = Arry(2).ToString
                            ElseIf Arry.Length = 4 Then
                                PoItemRow("Operator") = Arry(0).ToString
                                PoItemRow("StoragePosition") = Arry(1).ToString
                                PoItemRow("BoxDesc") = Arry(2).ToString
                                PoItemRow("ItemText") = Arry(3).ToString
                            End If
                        End If

                    Else
                        Dim TolQty As Decimal
                        PoItemRow("QtyBaseUOM") = PODetails.Tables("po_line").Rows(i)("BALANCE_QTY") * PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE")
                        TolQty = PODetails.Tables("po_line").Rows(i)("SHIP_TO_QUANTITY") * PODetails.Tables("po_line").Rows(i)("QTY_RCV_TOLERANCE") * 0.01
                        PoItemRow("GRQty") = (PODetails.Tables("po_line").Rows(i)("BALANCE_QTY") + TolQty) * PODetails.Tables("po_line").Rows(i)("CONVERSION_RATE")
                    End If

                    If PODetails.Tables("po_line").Rows(i)("SHELF_LIFE_DAYS") Is DBNull.Value Then PODetails.Tables("po_line").Rows(i)("SHELF_LIFE_DAYS") = 0
                    If PODetails.Tables("po_line").Rows(i)("MIN_SHELF_LIFE_REM_DAYS") Is DBNull.Value Then PODetails.Tables("po_line").Rows(i)("MIN_SHELF_LIFE_REM_DAYS") = 0

                    'Only the Expiration Date control is turned on, then we will look at Shelf Life Days, otherwise, set ExpDays = 0
                    If PoItemRow("ExpControl") = "Y" Then
                        If CInt(PODetails.Tables("po_line").Rows(i)("SHELF_LIFE_DAYS")) > 0 Then
                            PoItemRow("SLED") = "X"
                            PoItemRow("ExpDays") = CInt(PODetails.Tables("po_line").Rows(i)("SHELF_LIFE_DAYS"))
                            PoItemRow("MinRemDays") = CInt(PODetails.Tables("po_line").Rows(i)("MIN_SHELF_LIFE_REM_DAYS"))
                        Else
                            'Check if Invalid Review Period After Receipt in PIM settings  ---  06/03/2015
                            If PODetails.Tables("po_line").Rows(i)("Review_Days") = -1 Then               'Invalid Review Period after receipt in PIM settings
                                If GRHeader.OrderType = 4 Then
                                    TempStr = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString & " with ShipmentNo " & OrderNo & "/" & OrderItem
                                Else
                                    TempStr = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString & " with PO " & OrderNo & "/" & OrderItem
                                End If
                                ErrorMsg = "Invalid Review Period After Receipt in PIM settings for Item " & TempStr
                                ErrorRow = ErrorTable.NewRow()
                                ErrorRow(0) = ErrorMsg
                                ErrorTable.Rows.Add(ErrorRow)
                                Exit Function
                            End If
                        End If
                    End If

                    If PODetails.Tables("po_line").Rows(i)("receiving_routing_id").ToString = "2" Then PoItemRow("StockType") = "Q"
                    If PoItemRow("RevControl") = "Y" Then
                        If PODetails.Tables("po_line").Rows(i)("po_item_reversion").ToString <> "" Then
                            PoItemRow("MaterialRevision") = PODetails.Tables("po_line").Rows(i)("po_item_reversion").ToString
                        Else
                            PoItemRow("MaterialRevision") = PODetails.Tables("po_line").Rows(i)("item_rev").ToString
                        End If
                    End If

                    PoItemRow("PredefinedSubInv") = PODetails.Tables("po_line").Rows(i)("DESTINATION_SUBINVENTORY").ToString
                    PoItemRow("PredefinedLocator") = PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString
                End If

                ' Record AdditionalText/StoragePosition/DeliveryType/Operator/Manufacturer/ManufacturerPN from upload excel file
                If GRHeader.Type <> 2 And Not POLists Is Nothing Then
                    DR = Nothing
                    If GRHeader.Type = 1 Then
                        DR = POLists.Tables(0).Select(" PONo = '" & OrderNo & "' and POLine = '" & OrderItem & "'")
                    ElseIf GRHeader.Type = 3 Then
                        DR = POLists.Tables(0).Select(" PONo = '" & OrderNo & "' and POLine = '" & eJitPOLine & "'")
                    Else
                        Dim ASNNo, ASNLine As String
                        ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                        ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString
                        DR = POLists.Tables(0).Select(" PONo = '" & ASNNo & "' and POLine = '" & ASNLine & "'")
                    End If
                    If DR.Length > 0 Then
                        PoItemRow("ItemText") = DR(0)("AdditionalText").ToString
                        PoItemRow("StoragePosition") = DR(0)("StoragePosition").ToString
                        PoItemRow("Operator") = DR(0)("Operator").ToString
                        PoItemRow("Manufacturer") = DR(0)("Manufacturer").ToString
                        PoItemRow("ManufacturerPN") = DR(0)("ManufacturerPN").ToString

                        'Set default value for DeliveryType if the Currency is CNY with FY & LD    09/27/2013
                        PoItemRow("DeliveryType") = DR(0)("DeliveryType").ToString
                        If DeliveryType <> "" Then PoItemRow("DeliveryType") = DeliveryType

                        If PoItemRow("ExpControl") = "Y" Then
                            'Read ManufacturedDate from upload excel file for Shelf Life Controlled Parts
                            If GRHeader.CLIDFlag = False Then                     'No Vendor CLID provided from upload file
                                If PoItemRow("SLED") = "X" AndAlso (Not DR(0)("ManufactureDate") Is DBNull.Value) Then
                                    PoItemRow("ProductionDate") = CDate(DR(0)("ManufactureDate"))
                                End If
                            End If

                            'If Vendor CLID provided from upload file 
                            If GRHeader.CLIDFlag = True Then
                                If Not DR(0)("ExpDate") Is DBNull.Value Then PoItemRow("ExpDate") = CDate(DR(0)("ExpDate"))
                                If Not DR(0)("ManufactureDate") Is DBNull.Value Then PoItemRow("ProductionDate") = CDate(DR(0)("ManufactureDate"))

                                'ArtesynPN in upload file is different from Oracle Item, give error message            -- 01/16/2019
                                If POLists.Tables(0).Columns.Count >= 21 Then
                                    Dim n As Integer = 0
                                    Dim diffFlag As Boolean = False
                                    Dim tmpPN As String = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                                    For n = 0 To DR.Length - 1
                                        Dim APN As String = DR(n)("ArtesynPN").ToString
                                        If APN <> "" AndAlso APN <> tmpPN Then
                                            If GRHeader.Type <= 3 Then
                                                TempStr = tmpPN & " with PO/Line " & OrderNo & "/" & OrderItem
                                            Else
                                                Dim ASNNo, ASNLine As String
                                                ASNNo = PODetails.Tables("po_line").Rows(i)("ASN_SHIPMENT_NUM").ToString
                                                ASNLine = PODetails.Tables("po_line").Rows(i)("ASN_LINE_NUM").ToString
                                                TempStr = tmpPN & " with ASN/Line " & ASNNo & "/" & ASNLine
                                            End If
                                            ErrorMsg = "Different ArtesynPN found for Item " & TempStr
                                            ErrorRow = ErrorTable.NewRow()
                                            ErrorRow(0) = ErrorMsg
                                            ErrorTable.Rows.Add(ErrorRow)
                                            diffFlag = True
                                        End If
                                    Next
                                    If diffFlag = True Then Exit Function
                                End If

                                'Check if the same Item has different ExpDate in the upload file, if yes, give error message            -- 09/16/2016
                                Dim Exdr() As DataRow
                                Exdr = POItemData.Select("MaterialNo = '" & PoItemRow("MaterialNo") & "'")
                                If Exdr.Length > 0 Then
                                    Dim n As Integer
                                    Dim diffFlag As Boolean = False
                                    For n = 0 To Exdr.Length - 1
                                        If Exdr(n)("ExpDate") Is DBNull.Value AndAlso PoItemRow("ExpDate") Is DBNull.Value Then
                                        ElseIf Not Exdr(n)("ExpDate") Is DBNull.Value AndAlso PoItemRow("ExpDate") Is DBNull.Value Then
                                            diffFlag = True
                                            Exit For
                                        ElseIf Exdr(n)("ExpDate") Is DBNull.Value AndAlso Not PoItemRow("ExpDate") Is DBNull.Value Then
                                            diffFlag = True
                                            Exit For
                                        ElseIf Not Exdr(n)("ExpDate") Is DBNull.Value AndAlso Not PoItemRow("ExpDate") Is DBNull.Value Then
                                            If Exdr(n)("ExpDate") <> PoItemRow("ExpDate") Then
                                                diffFlag = True
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If diffFlag = True Then
                                        TempStr = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString & " with PO " & OrderNo & "/" & OrderItem
                                        ErrorMsg = "Different ExpDate found for Item " & TempStr
                                        ErrorRow = ErrorTable.NewRow()
                                        ErrorRow(0) = ErrorMsg
                                        ErrorTable.Rows.Add(ErrorRow)
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If


                        Dim j, k As Integer
                        Dim UnitQty, SumQty, UnitsNo, TotalRecQty As Decimal

                        SumQty = 0
                        TotalRecQty = 0
                        For j = 0 To DR.Length - 1                    'Record DateCode/LotNo/COO/UnitQty/NoOfPackage	
                            ' Record this line if it exists in the Upload file  -- 01/16/2019
                            DR(j)("Flag") = "1"
                            DR(j).AcceptChanges()

                            'Get COO Lists and check if the COO is valid or not when the COO is not blank
                            Dim COO As String = DR(j)("COO").ToString.Trim
                            If COO <> "" Then
                                If COOLists Is Nothing Then
                                    COOLists = New DataSet
                                    COOLists = GetCOOLists()
                                End If

                                If COOLists Is Nothing OrElse COOLists.Tables.Count = 0 Then
                                ElseIf COOLists.Tables(0).Rows.Count > 0 Then
                                    Dim myDR() As DataRow = Nothing
                                    myDR = COOLists.Tables(0).Select("territory_code = '" & COO & "'")
                                    If myDR.Length = 0 Then
                                        TempStr = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString & " with PO " & OrderNo & "/" & OrderItem
                                        ErrorMsg = "Invalid COO: " & COO & " for Item " & TempStr
                                        ErrorRow = ErrorTable.NewRow()
                                        ErrorRow("ErrorMsg") = ErrorMsg
                                        ErrorTable.Rows.Add(ErrorRow)
                                    End If
                                End If

                            End If

                            UnitQty = Convert.ToDecimal(DR(j)("UnitQty"))
                            UnitsNo = CInt(DR(j)("NoOfPackage"))
                            If GRHeader.CLIDFlag = True Then UnitsNo = 1 'VPNCLID provided by User           --  4/25/2016

                            SumQty = UnitQty * UnitsNo
                            TotalRecQty = TotalRecQty + SumQty

                            For k = 1 To UnitsNo
                                TraceabilityDataRow = TraceabilityData.NewRow()
                                TraceabilityDataRow("POItem") = OrderItem
                                TraceabilityDataRow("DateCode") = DR(j)("DateCode").ToString.Trim
                                TraceabilityDataRow("LotNo") = DR(j)("LotNo").ToString.Trim
                                TraceabilityDataRow("COO") = DR(j)("COO").ToString.Trim
                                TraceabilityDataRow("Qty") = String.Format("{0:F5}", UnitQty)
                                TraceabilityDataRow("LID") = ""

                                'VPNCLID provided by User           --  4/25/2016
                                If GRHeader.CLIDFlag = True Then TraceabilityDataRow("LID") = DR(j)("CLID").ToString.Trim

                                TraceabilityDataRow("PONo") = OrderNo
                                TraceabilityDataRow("POLineID") = POLineID
                                TraceabilityDataRow("MatNo") = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                                TraceabilityDataRow("RefCLID") = ""                         'Make sure it is not dbnull.value
                                TraceabilityDataRow("BoxID") = ""
                                TraceabilityDataRow("SubInventory") = DBNull.Value
                                TraceabilityDataRow("Locator") = DBNull.Value
                                TraceabilityDataRow("PredefinedSubInv") = PODetails.Tables("po_line").Rows(i)("DESTINATION_SUBINVENTORY").ToString
                                TraceabilityDataRow("PredefinedLocator") = PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString

                                TraceabilityData.Rows.Add(TraceabilityDataRow)
                            Next

                        Next

                        PoItemRow("EntryQty") = String.Format("{0:F5}", TotalRecQty)
                    End If
                End If

                'Get TraceabilityLevel and AddlData according to Commodity Code        'Yudy 04/30/2010
                PoItemRow("Traceable") = "N"
                If GRHeader.Type <> 5 Then
                    'Dim TraceLevel As MGTraceData
                    'Dim CommCode As String = Microsoft.VisualBasic.Left(PoItemRow("MaterialNo"), 3)
                    'TraceLevel = GetMGTraceData(CommCode)
                    'PoItemRow("Traceable") = TraceLevel.TraceabilityLevel
                    PoItemRow("Traceable") = GetMGTraceLevel(PoItemRow("MaterialNo"))
                End If

                If GRHeader.OrderType = 4 AndAlso FoundIRData = True Then
                    Dim MaterialNo, RTLot As String
                    MaterialNo = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                    RTLot = FixNull(PODetails.Tables("po_line").Rows(i)("RT_LOT_NUMBER"))
                    DR = Nothing
                    DR = myCLIDs.Tables(0).Select(" MaterialNo = '" & MaterialNo & "' and ( RTLot = '" & RTLot & "'or RTLot IS NULL) ")
                    If DR.Length > 0 Then
                        'CLID, MaterialNo, QtyBaseUOM, BaseUOM, DateCode, LotNo, CountryOfOrigin, RTLot, ExpDate, ProdDate, RoHS, 
                        'MatSuffix1, MatSuffix2, MatSuffix3, Manufacturer, ManufacturerPN, QMLStatus, AddlData, Stemp, MSL
                        PoItemRow("Manufacturer") = DR(0)("Manufacturer").ToString.Trim
                        PoItemRow("ManufacturerPN") = DR(0)("ManufacturerPN").ToString.Trim
                        PoItemRow("QMLStatus") = DR(0)("QMLStatus").ToString.Trim
                        PoItemRow("MatSuffix1") = DR(0)("MatSuffix1").ToString.Trim
                        PoItemRow("MatSuffix2") = DR(0)("MatSuffix2").ToString.Trim
                        PoItemRow("MatSuffix3") = DR(0)("MatSuffix3").ToString.Trim
                        PoItemRow("RoHS") = DR(0)("RoHS").ToString.Trim
                        PoItemRow("AddlData") = DR(0)("AddlData").ToString.Trim
                        PoItemRow("Stemp") = DR(0)("Stemp").ToString.Trim
                        PoItemRow("MSL") = DR(0)("MSL").ToString.Trim

                        'Read ManufacturedDate / ExpDate for Shelf Life Controlled Parts
                        If PoItemRow("ExpControl") = "Y" Then
                            If Not DR(0)("ExpDate") Is DBNull.Value Then         'ExpDate found from Source,Set SLED = ""
                                If PoItemRow("SLED") = "X" Then PoItemRow("SLED") = ""

                                PoItemRow("ExpDate") = CDate(DR(0)("ExpDate"))
                                If Not DR(0)("ProdDate") Is DBNull.Value Then PoItemRow("ProductionDate") = CDate(DR(0)("ProdDate"))
                            End If
                        End If

                        Dim j As Integer
                        Dim TotalRecQty As Decimal = 0
                        Dim OpenIRQty As Decimal = PoItemRow("QtyBaseUOM")

                        For j = 0 To DR.Length - 1
                            If DR(j)("QtyBaseUOM") = 0 AndAlso DR(j)("Status") = 0 Then
                            Else
                                If TotalRecQty = PoItemRow("QtyBaseUOM") OrElse OpenIRQty = 0 Then Exit For

                                Dim LID, RefCLID As String
                                Dim CLIDQty, RestQty As Decimal
                                RestQty = DR(j)("QtyBaseUOM")
                                If OpenIRQty > RestQty Then
                                    CLIDQty = RestQty
                                Else
                                    CLIDQty = OpenIRQty
                                End If

                                If DR(j)("Status") = 0 And OpenIRQty >= RestQty Then
                                    LID = DR(j)("CLID")
                                    RefCLID = ""
                                Else
                                    LID = ""
                                    RefCLID = DR(j)("CLID")
                                End If
                                OpenIRQty = OpenIRQty - CLIDQty
                                RestQty = RestQty - CLIDQty

                                DR(j)("QtyBaseUOM") = RestQty
                                If RestQty > 0 Then
                                    DR(j)("Status") = 1
                                Else
                                    DR(j)("Status") = 0
                                End If
                                DR(j).AcceptChanges()

                                TraceabilityDataRow = TraceabilityData.NewRow()
                                TraceabilityDataRow("POItem") = OrderItem
                                TraceabilityDataRow("DateCode") = DR(j)("DateCode").ToString.Trim
                                TraceabilityDataRow("LotNo") = DR(j)("LotNo").ToString.Trim
                                TraceabilityDataRow("COO") = DR(j)("CountryOfOrigin").ToString.Trim
                                TraceabilityDataRow("Qty") = String.Format("{0:F5}", CLIDQty)   'DR(j)("QtyBaseUOM")
                                TraceabilityDataRow("LID") = LID                                'DR(j)("CLID")
                                TraceabilityDataRow("PONo") = OrderNo
                                TraceabilityDataRow("POLineID") = POLineID
                                TraceabilityDataRow("MatNo") = PODetails.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                                TraceabilityDataRow("RefCLID") = RefCLID      '""            'Make sure it is not dbnull.value
                                TraceabilityDataRow("BoxID") = ""
                                TraceabilityDataRow("SubInventory") = DBNull.Value
                                TraceabilityDataRow("Locator") = DBNull.Value
                                TraceabilityDataRow("PredefinedSubInv") = PODetails.Tables("po_line").Rows(i)("DESTINATION_SUBINVENTORY").ToString
                                TraceabilityDataRow("PredefinedLocator") = PODetails.Tables("po_line").Rows(i)("DESTINATION_LOCATOR").ToString

                                'Save TransactionID for later search use only for ReceiptType = 9
                                'If GRHeader.Type = 9 Then TraceabilityDataRow("Status") = POLineID

                                TraceabilityData.Rows.Add(TraceabilityDataRow)
                                TotalRecQty = TotalRecQty + CLIDQty
                            End If
                        Next
                        PoItemRow("EntryQty") = String.Format("{0:F5}", TotalRecQty)
                    End If

                End If

                POItemData.Rows.Add(PoItemRow)
                PODetails.Tables("po_line").AcceptChanges()
            Next
            PODetails.AcceptChanges()
            If ErrorTable.Rows.Count > 0 Then Exit Function

            GetItemsFromERP.Merge(PODetails.Tables("po_header"))
            GetItemsFromERP.Merge(PODetails.Tables("po_line"))

            ' Read Shipment Data if there exists BoxID/PalletID for OSP PO Upload
            If GRHeader.Type = 2 AndAlso Not POLists Is Nothing Then
                AddShipData(LoginData, POLists, GetItemsFromERP)
            End If

            'Upload PO records must be the same with Oracle records which extracted, otherwise, give error message   -- 01/16/2019
            If Not POLists Is Nothing Then
                DR = Nothing
                DR = POLists.Tables(0).Select("Flag = '' ")
                If DR.Length > 0 Then
                    For i = 0 To DR.Length - 1
                        If GRHeader.Type = 6 Or GRHeader.Type = 7 Then
                            ErrorMsg = "Invalid ASN No# " & DR(i)("PONo").ToString & "/" & DR(i)("POLine").ToString
                        Else
                            ErrorMsg = "Invalid PO " & DR(i)("PONo").ToString & "/" & DR(i)("POLine").ToString
                        End If
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow("ErrorMsg") = ErrorMsg
                        ErrorTable.Rows.Add(ErrorRow)
                    Next
                End If
            End If
            If ErrorTable.Rows.Count > 0 Then Exit Function


            If GRHeader.Type <> 2 AndAlso Not POLists Is Nothing Then
                ' ''Sort PO batch Data by Ascending according to column  ItemText / AdditionalText   --- 01/25/2016      -- Cancelled on 2/17/2016
                'Dim dtBatchPO As DataTable = New DataTable
                ''Dim SortColName As String = POItemData.Columns(10).ColumnName      'ItemText / AdditionalText
                'SortColName = SortColName & " ASC"
                'POItemData.DefaultView.Sort = SortColName
                'dtBatchPO = POItemData.DefaultView.ToTable()
                'POItemData.Clear()
                'POItemData.Merge(dtBatchPO)

                'Read COO Lists if there exists COO for Normal PO upload
                If COOLists Is Nothing OrElse COOLists.Tables.Count = 0 Then
                ElseIf COOLists.Tables(0).Rows.Count > 0 Then
                    GetItemsFromERP.Merge(COOLists.Tables("COOData"))
                End If
            End If


            'Read problematic D/C & L/C from eTrace table    01/18/2017
            If GRHeader.OrderType = 1 OrElse GRHeader.OrderType = 3 OrElse _
                 GRHeader.OrderType = 6 OrElse GRHeader.OrderType = 7 Then
                Dim Sqlstr As String
                Dim da As DataAccess = GetDataAccess()
                Dim dsDCLN As DataSet = New DataSet
                Sqlstr = String.Format("Select * from T_DateCodeLotNo ")
                dsDCLN = da.ExecuteDataSet(Sqlstr, "DateCodeLN")
                GetItemsFromERP.Merge(dsDCLN.Tables(0))
            End If


        Catch ex As Exception
            ErrorLogging("Receiving-GetItemsFromERP", LoginData.User.ToUpper, "OrderNo: " & GRHeader.OrderNo & ", " & ex.Message & ex.Source, "E")
            ErrorMsg = "Invalid PO " & GRHeader.OrderNo
            If GRHeader.Type = 6 Or GRHeader.Type = 7 Then ErrorMsg = "Invalid ASN No# " & GRHeader.OrderNo
            If GRHeader.Type = 9 Then ErrorMsg = "Invalid Receipt Number " & GRHeader.OrderNo
            ErrorRow = ErrorTable.NewRow()
            ErrorRow("ErrorMsg") = ErrorMsg
            ErrorTable.Rows.Add(ErrorRow)
        End Try

    End Function

    Private Function AddShipData(ByVal LoginData As ERPLogin, ByVal POLists As DataSet, ByRef Items As DataSet) As DataSet

        Dim i As Integer
        Dim myDataRow As Data.DataRow
        Dim ErrorRow As Data.DataRow

        Dim DR() As DataRow = Nothing
        Dim dsShipData As DataSet = Nothing

        Try
            DR = POLists.Tables(0).Select(" LotNo <> '' ")
            If DR.Length > 0 Then
                Dim OSPDJ As String = ""
                For i = 0 To Items.Tables("POItemData").Rows.Count - 1
                    If OSPDJ = "" Then
                        OSPDJ = Items.Tables("POItemData").Rows(i)("ReasonCode").ToString
                    Else
                        OSPDJ = OSPDJ & "," & Items.Tables("POItemData").Rows(i)("ReasonCode").ToString
                    End If
                Next

                Try
                    dsShipData = ReadShipmentData(LoginData, OSPDJ)
                Catch ex As Exception
                    ErrorLogging("Receiving-AddShipData", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
                    dsShipData = Nothing
                End Try

                If dsShipData Is Nothing OrElse dsShipData.Tables.Count = 0 OrElse dsShipData.Tables("ShipData").Rows.Count = 0 Then
                    ErrorRow = Items.Tables("ErrorTable").NewRow()
                    ErrorRow("ErrorMsg") = "^REC-78@"
                    Items.Tables("ErrorTable").Rows.Add(ErrorRow)
                    Return Items
                    Exit Function
                End If

                Items.Merge(dsShipData.Tables("CLIDData"))
                Items.Merge(dsShipData.Tables("ShipData"))

            End If

            For i = 0 To Items.Tables("POItemData").Rows.Count - 1
                Dim OrderNo, OrderItem As String
                OrderNo = Items.Tables("POItemData").Rows(i)("PONo").ToString
                OrderItem = Items.Tables("POItemData").Rows(i)("POItem").ToString

                DR = Nothing
                DR = POLists.Tables(0).Select(" PONo = '" & OrderNo & "' and POLine = '" & OrderItem & "'")
                If DR.Length > 0 Then
                    Items.Tables("POItemData").Rows(i)("ItemText") = DR(0)("AdditionalText").ToString
                    Items.Tables("POItemData").Rows(i)("Operator") = DR(0)("Operator").ToString

                    Dim j, k As Integer
                    Dim UnitQty, SumQty, UnitsNo, TotalRecQty As Decimal

                    SumQty = 0
                    TotalRecQty = 0
                    For j = 0 To DR.Length - 1                    'Record DateCode / BoxID/PalletID / UnitQty / NoOfPackage
                        ' Record this PO Line if it exists in the Upload file  -- 01/16/2019
                        DR(j)("Flag") = "1"
                        DR(j).AcceptChanges()

                        Dim LabelID As String = DR(j)("LotNo").ToString.Trim
                        Dim DateCode As String = DR(j)("DateCode").ToString.Trim

                        UnitQty = Convert.ToDecimal(DR(j)("UnitQty"))
                        UnitsNo = CInt(DR(j)("NoOfPackage"))

                        If UnitQty > 0 AndAlso UnitsNo > 0 Then
                            SumQty = UnitQty * UnitsNo
                            TotalRecQty = TotalRecQty + SumQty

                            For k = 1 To UnitsNo
                                myDataRow = Items.Tables("TraceabilityData").NewRow()
                                myDataRow("POItem") = OrderItem
                                myDataRow("DateCode") = DateCode
                                myDataRow("LotNo") = ""
                                myDataRow("COO") = ""
                                myDataRow("Qty") = String.Format("{0:F5}", UnitQty)
                                myDataRow("LID") = ""
                                myDataRow("PONo") = OrderNo
                                myDataRow("MatNo") = Items.Tables("POItemData").Rows(i)("MaterialNo").ToString
                                myDataRow("RefCLID") = ""                         'Make sure it is not dbnull.value
                                myDataRow("BoxID") = ""
                                myDataRow("SubInventory") = Items.Tables("POItemData").Rows(i)("SubInventory").ToString
                                myDataRow("Locator") = Items.Tables("POItemData").Rows(i)("Locator").ToString
                                myDataRow("PredefinedSubInv") = DBNull.Value
                                myDataRow("PredefinedLocator") = DBNull.Value

                                Items.Tables("TraceabilityData").Rows.Add(myDataRow)
                            Next

                        Else

                            Dim Found As Boolean = False
                            Dim LabelType As String = ""
                            If Microsoft.VisualBasic.Left(LabelID, 1) = "B" Then                      'CartonID
                                LabelType = "CLID"
                            ElseIf Microsoft.VisualBasic.Left(LabelID, 1) = "P" Then                  'PalletID
                                LabelType = "BoxID"
                            End If

                            Dim ValidCLID As String = ""
                            Dim myCLIDs As DataTable = New DataTable

                            myCLIDs = Items.Tables("TraceabilityData")
                            ValidCLID = Validate_ShipData(LabelID, myCLIDs, dsShipData)

                            If ValidCLID <> "" Then
                                ErrorRow = Items.Tables("ErrorTable").NewRow()
                                ErrorRow("ErrorMsg") = LabelID & ": " & ValidCLID
                                Items.Tables("ErrorTable").Rows.Add(ErrorRow)
                                Return Items
                                Exit Function
                            End If


                            Dim Shipdr() As DataRow = Nothing
                            If LabelType = "CLID" Then
                                Shipdr = dsShipData.Tables("ShipData").Select(" CLID = '" & LabelID & "'")
                            Else
                                Shipdr = dsShipData.Tables("ShipData").Select(" BoxID = '" & LabelID & "'")
                            End If
                            If Shipdr.Length > 0 Then
                                For k = 0 To Shipdr.Length - 1
                                    Dim Flag As Boolean = False
                                    Dim CLIDdr() As DataRow = Nothing
                                    If LabelType = "BoxID" Then           'Check if the LabelID already entered
                                        CLIDdr = Items.Tables("TraceabilityData").Select(" LID = '" & Shipdr(k)("CLID") & "'")
                                        If CLIDdr.Length = 0 Then
                                            Flag = True
                                        Else
                                            CLIDdr(0)("BoxID") = LabelID
                                            CLIDdr(0)("RefCLID") = "BoxID"
                                            CLIDdr(0).AcceptChanges()
                                            Found = True
                                        End If
                                    Else
                                        Flag = True
                                    End If

                                    If Flag = True Then
                                        myDataRow = Items.Tables("TraceabilityData").NewRow()
                                        myDataRow("POItem") = OrderItem
                                        myDataRow("DateCode") = DateCode
                                        myDataRow("LotNo") = Shipdr(k)("CLID")
                                        myDataRow("COO") = ""
                                        myDataRow("Qty") = String.Format("{0:F5}", Shipdr(k)("QtyBaseUOM"))     'Shipdr(k)("QtyBaseUOM")   
                                        myDataRow("LID") = Shipdr(k)("CLID")
                                        myDataRow("PONo") = OrderNo
                                        myDataRow("MatNo") = Items.Tables("POItemData").Rows(i)("MaterialNo").ToString

                                        If LabelType = "BoxID" Then
                                            myDataRow("BoxID") = Shipdr(k)("BoxID").ToString
                                            myDataRow("RefCLID") = "BoxID"                         'temporary field, identify the input value is CLID or BoxID
                                        Else
                                            myDataRow("BoxID") = ""
                                            myDataRow("RefCLID") = "CLID"                          'temporary field, identify the input value is CLID or BoxID
                                        End If
                                        myDataRow("SubInventory") = Items.Tables("POItemData").Rows(i)("SubInventory").ToString
                                        myDataRow("Locator") = Items.Tables("POItemData").Rows(i)("Locator").ToString
                                        myDataRow("PredefinedSubInv") = DBNull.Value
                                        myDataRow("PredefinedLocator") = DBNull.Value
                                        Items.Tables("TraceabilityData").Rows.Add(myDataRow)

                                        TotalRecQty = TotalRecQty + Shipdr(k)("QtyBaseUOM")
                                    End If
                                Next
                            End If
                        End If
                    Next

                    Items.Tables("POItemData").Rows(i)("EntryQty") = String.Format("{0:F5}", TotalRecQty)
                    Items.Tables("POItemData").Rows(i).AcceptChanges()
                End If
            Next
            Items.AcceptChanges()

            Return Items

        Catch ex As Exception
            ErrorLogging("Receiving-AddShipData", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            ErrorRow = Items.Tables("ErrorTable").NewRow()
            ErrorRow("ErrorMsg") = "^REC-78@"
            Items.Tables("ErrorTable").Rows.Add(ErrorRow)
            Return Items
        End Try

    End Function

    Private Function Validate_ShipData(ByVal LabelID As String, ByVal CLIDs As DataTable, ByVal dsShipData As DataSet) As String
        Validate_ShipData = ""

        Try
            Dim LabelType As String = ""
            If Microsoft.VisualBasic.Left(LabelID, 1) = "B" Then                      'CartonID
                LabelType = "CLID"
            ElseIf Microsoft.VisualBasic.Left(LabelID, 1) = "P" Then                  'PalletID
                LabelType = "BoxID"
            End If

            Dim DR() As DataRow = Nothing
            If CLIDs.Rows.Count > 0 Then
                If LabelType = "CLID" Then
                    DR = CLIDs.Select(" LID = '" & LabelID & "'")
                Else
                    DR = CLIDs.Select(" BoxID = '" & LabelID & "'")
                End If
                If DR.Length > 0 Then
                    If LabelType = "BoxID" And DR(0)("RefCLID") = "CLID" Then
                    Else
                        Validate_ShipData = "^REC-79@"
                        Exit Function
                    End If
                End If
            End If

            DR = Nothing
            If dsShipData.Tables("CLIDData").Rows.Count > 0 Then
                If LabelType = "CLID" Then
                    DR = dsShipData.Tables("CLIDData").Select(" CLID = '" & LabelID & "'")
                Else
                    DR = dsShipData.Tables("CLIDData").Select(" BoxID = '" & LabelID & "'")
                End If
                If DR.Length > 0 Then
                    Validate_ShipData = "^REC-79@"
                    Exit Function
                End If
            End If

            DR = Nothing
            If LabelType = "CLID" Then
                DR = dsShipData.Tables("ShipData").Select(" CLID = '" & LabelID & "'")
            Else
                DR = dsShipData.Tables("ShipData").Select(" BoxID = '" & LabelID & "'")
            End If
            If DR.Length = 0 Then
                Validate_ShipData = "^REC-78@"
            End If

        Catch ex As Exception
            ErrorLogging("Receiving-Validate_ShipData", "", ex.Message & ex.Source, "E")
            Validate_ShipData = "^REC-78@"
        End Try

    End Function

    Public Function ProcessMatMovement(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet, ByVal PrintLabels As Boolean, ByVal LabelPrinter As String) As CreateGRResponse
        Try
            Select Case Header.Type
                Case 8
                    ProcessMatMovement = CancelGR(LoginData, Header, Items)
                Case 11
                    ProcessMatMovement = VerifyRTCLID(LoginData, Header)
                Case 12
                    ProcessMatMovement = SaveDHLData(LoginData, Header, Items)
                Case Else
                    ProcessMatMovement = PostGR(LoginData, Header, Items, PrintLabels, LabelPrinter)
            End Select
        Catch ex As Exception
            ErrorLogging("Receiving-ProcessMatMovement", LoginData.User.ToUpper, "OrderNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
        End Try

    End Function

    Public Function GetEJitData(ByVal LoginData As ERPLogin, ByVal ds As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim sda As SqlClient.SqlDataAdapter = da.Sda_Insert()
            Dim RTNo As String = ds.Tables(0).Rows(0)("RTNo").ToString

            Try
                sda.InsertCommand.CommandType = CommandType.StoredProcedure
                sda.InsertCommand.CommandText = "ora_get_ejit_receipted_data"
                sda.InsertCommand.CommandTimeout = TimeOut_M5

                sda.InsertCommand.Parameters.Add("@org_id", SqlDbType.Int).Value = CInt(LoginData.OrgID)
                sda.InsertCommand.Parameters.Add("@trn_id", SqlDbType.Int)
                sda.InsertCommand.Parameters.Add("@receipt_type", SqlDbType.VarChar, 10)
                sda.InsertCommand.Parameters.Add("@destination_subinventory", SqlDbType.VarChar, 50)
                sda.InsertCommand.Parameters.Add("@destination_locator", SqlDbType.VarChar, 50)
                sda.InsertCommand.Parameters.Add("@ejit_id", SqlDbType.VarChar, 50)
                sda.InsertCommand.Parameters.Add("@rt_lot", SqlDbType.VarChar, 50)
                sda.InsertCommand.Parameters.Add("@exp_date", SqlDbType.VarChar, 50)
                sda.InsertCommand.Parameters("@destination_subinventory").Direction = ParameterDirection.Output
                sda.InsertCommand.Parameters("@destination_locator").Direction = ParameterDirection.Output
                sda.InsertCommand.Parameters("@ejit_id").Direction = ParameterDirection.Output
                sda.InsertCommand.Parameters("@rt_lot").Direction = ParameterDirection.Output
                sda.InsertCommand.Parameters("@exp_date").Direction = ParameterDirection.Output

                sda.InsertCommand.Parameters("@trn_id").SourceColumn = "TranID"
                sda.InsertCommand.Parameters("@receipt_type").SourceColumn = "RecType"
                sda.InsertCommand.Parameters("@destination_subinventory").SourceColumn = "SubInv"
                sda.InsertCommand.Parameters("@destination_locator").SourceColumn = "Locator"
                sda.InsertCommand.Parameters("@ejit_id").SourceColumn = "eJITID"
                sda.InsertCommand.Parameters("@rt_lot").SourceColumn = "RTLot"
                sda.InsertCommand.Parameters("@exp_date").SourceColumn = "ExpDate"

                sda.InsertCommand.Connection.Open()
                sda.Update(ds.Tables(0))
                sda.InsertCommand.Connection.Close()
                Return ds

            Catch ex As Exception
                ErrorLogging("Receiving-GetEJitData", LoginData.User, "RTNo: " & RTNo & "; " & ex.Message & ex.Source, "E")
                Return ds
            Finally
                If sda.InsertCommand.Connection.State <> ConnectionState.Closed Then sda.InsertCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function ReadOSPAssemblyUOM(ByVal LoginData As ERPLogin, ByVal dsOSPData As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i As Integer
                Dim SqlStr As String
                Dim OSPDJStr As String = ""

                For i = 0 To dsOSPData.Tables("po_line").Rows.Count - 1
                    If OSPDJStr = "" Then
                        OSPDJStr = dsOSPData.Tables("po_line").Rows(i)("WIP_ENTITY_NAME").ToString
                    Else
                        OSPDJStr = OSPDJStr & "," & dsOSPData.Tables("po_line").Rows(i)("WIP_ENTITY_NAME").ToString
                    End If
                Next

                Try
                    Dim ds As New DataSet()
                    SqlStr = String.Format("exec dbo.sp_DJmodel_info '{0}','{1}'", LoginData.OrgCode, OSPDJStr)
                    ds = da.ExecuteDataSet(SqlStr, "DJModel")
                    Return ds

                Catch ex As Exception
                    ErrorLogging("Receiving-Call:sp_DJmodel_info", LoginData.User, ex.Message & ex.Source, "E")
                    Return Nothing
                End Try

            Catch ex As Exception
                ErrorLogging("Receiving-ReadOSPAssemblyUOM", LoginData.User, ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Private Function PostGR(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet, ByVal PrintLabels As Boolean, ByVal LabelPrinter As String) As CreateGRResponse
        PostGR = New CreateGRResponse

        Dim PODetails As New DataSet
        Dim i, j As Integer
        Dim DR() As DataRow = Nothing
        Dim GRNo As String = ""
        Dim PostDate As Date             'Get Date and Time

        Dim IRLabels As DataSet = New DataSet
        Dim CLIDs As DataTable
        Dim IRDataColumn As DataColumn
        Dim KeyColumn(0) As DataColumn
        Dim IRDataRow As Data.DataRow

        CLIDs = New Data.DataTable("CLIDs")
        IRDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CLIDs.Columns.Add(IRDataColumn)
        KeyColumn(0) = IRDataColumn
        CLIDs.PrimaryKey = KeyColumn      'This is to ensure we collect IR lables only once
        IRLabels.Tables.Add(CLIDs)

        Dim SqlStr As String
        Dim RecConfig As String = ""
        Dim Sqlcmd As SqlClient.SqlCommand
        Dim dtDashB As DataTable = New DataTable
        dtDashB.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
        dtDashB.Columns.Add(New Data.DataColumn("Packages", System.Type.GetType("System.String")))
        dtDashB.Columns.Add(New Data.DataColumn("RecQty", System.Type.GetType("System.Decimal")))
        dtDashB.Columns.Add(New Data.DataColumn("RecRouting", System.Type.GetType("System.String")))
        dtDashB.Columns.Add(New Data.DataColumn("PalletID", System.Type.GetType("System.String")))

        Try

            Dim EJitRTFlag As Boolean = False
            Dim dsEJit As DataSet = New DataSet
            Dim dtEJit As DataTable = New DataTable

            dtEJit.Columns.Add(New Data.DataColumn("RTNo", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("TranID", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("RecType", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("SubInv", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("Locator", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("eJITID", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("RTLot", System.Type.GetType("System.String")))
            dtEJit.Columns.Add(New Data.DataColumn("ExpDate", System.Type.GetType("System.String")))
            dsEJit.Tables.Add(dtEJit)

            PODetails.Merge(Items.Tables("po_header"))
            PODetails.Merge(Items.Tables("po_line"))

            PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("REC_QTY", System.Type.GetType("System.Decimal")))
            PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("REC_PRIMARY_QTY", System.Type.GetType("System.Decimal")))
            PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("AML_STATUS", System.Type.GetType("System.String")))

            If Header.Type = 6 Or Header.Type = 7 Then
                PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("lot_num", System.Type.GetType("System.String")))
                PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("lot_expiration_date", System.Type.GetType("System.DateTime")))
            Else
                'Package needs to Read COO data from Oracle for ASN Receipt, so the column already there for ASN Type, and just create for NON-ASN Type
                PODetails.Tables("po_line").Columns.Add(New Data.DataColumn("COO", System.Type.GetType("System.String")))
            End If

            Dim GroupID As Integer
            Dim HeaderInterfaceID As Integer
            Dim myRcvID As RcvID = New RcvID

            'Record PO Currency for adding "CN" suffix code into Oracle Lot for CNY purchase items  10/16/2014
            Header.POCurrency = PODetails.Tables("po_header").Rows(0)("currency_code").ToString
            Header.PostDate = Date.Today

            If Header.Type = 9 Then                            'CLID creation based on Oracle Receipt
                GRNo = Header.GRNo
                'Bug fix to make sure RTNo is not blank in table T_RTSlip   10/21/2014
                If FixNull(GRNo) = "" Then GRNo = Header.OrderNo
                PostGR.GRNo = GRNo
                PostGR.GRStatus = "Y"                             'Receipt No created 
                PostGR.GRMessage = "Receipt full completed"
                PostDate = CDate(PODetails.Tables("po_header").Rows(0)("receipt_date"))
            Else
                'HeaderInterfaceID = Get_HeaderInterfaceID()

                'Get HeaderInterfaceID and GroupID same as Oracle form   Yudy 07/06/2012 
                myRcvID = Get_HeaderGroupID()
                GroupID = myRcvID.GroupID
                HeaderInterfaceID = myRcvID.HeaderID

                'Fill in PO Header
                PODetails.Tables("po_header").Rows(0)("USER_ID") = LoginData.UserID
                PODetails.Tables("po_header").Rows(0)("waybill_airbill_num") = Header.DeliveryNote
                PODetails.Tables("po_header").Rows(0)("bill_of_lading") = Header.BillOfLading
                PODetails.Tables("po_header").Rows(0)("comments") = Header.HeaderText
                PODetails.Tables("po_header").Rows(0)("packing_slip") = Header.InvoiceNo
                PODetails.Tables("po_header").Rows(0)("header_interface_id") = HeaderInterfaceID
                PODetails.Tables("po_header").AcceptChanges()
                PODetails.Tables("po_header").Rows(0).SetAdded()
            End If

            Dim ProdDate As DateTime
            Dim ExpiryDate As DateTime
            Dim ExpDays As Decimal
            Dim DRLine() As DataRow
            Dim GRResult As New DataSet

            Dim dsPOData As DataSet = New DataSet
            dsPOData = PODetails.Copy

            If Header.Type <> 9 Then dsPOData.Tables("po_line").Clear()
            'Dim myLotNumber As Long = CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))    'Generate LotNumber with one date format 

            For i = 0 To Items.Tables(0).Rows.Count - 1
                If Items.Tables(0).Rows(i)("EntryQty") Is DBNull.Value OrElse Items.Tables(0).Rows(i)("EntryQty") = "" Then
                    Items.Tables(0).Rows(i)("EntryQty") = 0

                ElseIf Convert.ToDecimal(Items.Tables(0).Rows(i)("EntryQty")) > 0 Then
                    ExpDays = 0
                    ProdDate = New DateTime
                    ExpiryDate = New DateTime

                    If Items.Tables(0).Rows(i)("MaterialNo") Is DBNull.Value Then Items.Tables(0).Rows(i)("MaterialNo") = ""
                    If Items.Tables(0).Rows(i)("MaterialRevision") Is DBNull.Value Then Items.Tables(0).Rows(i)("MaterialRevision") = ""
                    If Items.Tables(0).Rows(i)("StockType") Is DBNull.Value Then Items.Tables(0).Rows(i)("StockType") = ""
                    If Items.Tables(0).Rows(i)("POItem") Is DBNull.Value Then Items.Tables(0).Rows(i)("POItem") = 0
                    If Items.Tables(0).Rows(i)("StoragePosition") Is DBNull.Value Then Items.Tables(0).Rows(i)("StoragePosition") = ""
                    If Items.Tables(0).Rows(i)("DeliveryType") Is DBNull.Value Then Items.Tables(0).Rows(i)("DeliveryType") = ""
                    'If Items.Tables(0).Rows(i)("Operator") Is DBNull.Value Then Items.Tables(0).Rows(i)("Operator") = ""
                    If Items.Tables(0).Rows(i)("MatSuffix1") Is DBNull.Value Then Items.Tables(0).Rows(i)("MatSuffix1") = ""
                    If Items.Tables(0).Rows(i)("MatSuffix2") Is DBNull.Value Then Items.Tables(0).Rows(i)("MatSuffix2") = ""
                    If Items.Tables(0).Rows(i)("MatSuffix3") Is DBNull.Value Then Items.Tables(0).Rows(i)("MatSuffix3") = ""
                    If Items.Tables(0).Rows(i)("RoHS") Is DBNull.Value Then Items.Tables(0).Rows(i)("RoHS") = ""
                    If Items.Tables(0).Rows(i)("ReasonCode") Is DBNull.Value Then Items.Tables(0).Rows(i)("ReasonCode") = ""
                    If Items.Tables(0).Rows(i)("ItemText") Is DBNull.Value Then Items.Tables(0).Rows(i)("ItemText") = ""
                    If Items.Tables(0).Rows(i)("BoxDesc") Is DBNull.Value Then Items.Tables(0).Rows(i)("BoxDesc") = ""
                    If Items.Tables(0).Rows(i)("AddlData") Is DBNull.Value Then Items.Tables(0).Rows(i)("AddlData") = ""
                    If Items.Tables(0).Rows(i)("Stemp") Is DBNull.Value Then Items.Tables(0).Rows(i)("Stemp") = ""
                    If Items.Tables(0).Rows(i)("MSL") Is DBNull.Value Then Items.Tables(0).Rows(i)("MSL") = ""
                    If Items.Tables(0).Rows(i)("COC") Is DBNull.Value Then Items.Tables(0).Rows(i)("COC") = ""
                    If Items.Tables(0).Rows(i)("EER") Is DBNull.Value Then Items.Tables(0).Rows(i)("EER") = ""
                    If Items.Tables(0).Rows(i)("RTLot") Is DBNull.Value Then Items.Tables(0).Rows(i)("RTLot") = ""
                    If FixNull(Items.Tables(0).Rows(i)("Operator")) = "" Then Items.Tables(0).Rows(i)("Operator") = LoginData.User


                    'Caculate ExpDate / ProductionDate only if the Expiration Date Control for that item is turned on
                    If Items.Tables(0).Rows(i)("ExpControl") = "Y" Then
                        If Not Items.Tables(0).Rows(i)("ExpDays") Is DBNull.Value Then ExpDays = Items.Tables(0).Rows(i)("ExpDays")
                        If Not Items.Tables(0).Rows(i)("ProductionDate") Is DBNull.Value Then ProdDate = Items.Tables(0).Rows(i)("ProductionDate")
                        If Not Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then ExpiryDate = Items.Tables(0).Rows(i)("ExpDate")

                        If Items.Tables(0).Rows(i)("ProductionDate") Is DBNull.Value AndAlso Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                            If Header.OrderType <> 2 Then            'Non OSP PO Type
                                'For those Parts with Shelf Life Days = 0, first check if the Review Period after receipt in PIM is 0 or not  -- 06/03/2015
                                'If Review Period after receipt = 0, then ExpiryDate = Receiving date + 3 Years
                                'If Review Period after receipt <> 0, then ExpiryDate = Receiving date + Review Period after receipt
                                Dim myReviewDays As Integer = 0
                                DR = PODetails.Tables("po_line").Select(" POLineID = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                                If DR.Length > 0 Then myReviewDays = CInt(DR(0)("Review_Days"))
                                ProdDate = Date.Today                'Use Today instead of Header.PostDate, avoid ExpDate becomes 01/01/0004 which happend in PH 
                                If myReviewDays = 0 Then
                                    ExpiryDate = ProdDate.AddDays(365 * 3)
                                Else
                                    ExpiryDate = ProdDate.AddDays(myReviewDays)
                                End If
                                Items.Tables(0).Rows(i)("ProductionDate") = ProdDate
                                Items.Tables(0).Rows(i)("ExpDate") = ExpiryDate
                            End If
                        ElseIf Not Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then      'ExpDate not Null, calculate ProdDate automatically
                            If Header.OrderType <> 4 Then            'Non IR
                                ProdDate = ExpiryDate.AddDays(ExpDays * (-1))
                                'For vendor provided ExpDate, if TotalShelfLifeDays=0, then Use Today instead of ProdDate         -- 9/16/2016
                                If ExpDays = 0 Then ProdDate = Date.Today
                                Items.Tables(0).Rows(i)("ProductionDate") = ProdDate
                            End If
                        Else                                                                    'ProdDate not Null, calculate ExpDate automatically
                            ExpiryDate = ProdDate.AddDays(ExpDays)
                            Items.Tables(0).Rows(i)("ExpDate") = DateValue(ExpiryDate)
                        End If

                        'Valid ExpiryDate --4/11/2018
                        If DateDiff(DateInterval.Day, DateTime.Now, ExpiryDate, ) <= 0 Then
                            PostGR.GRNo = ""
                            PostGR.GRStatus = "E"
                            PostGR.GRMessage = "PO: " & Items.Tables(0).Rows(i)("PONo") & ",Line: " & Items.Tables(0).Rows(i)("POItem") & ",Item: " & Items.Tables(0).Rows(i)("MaterialNo") & ", " & "^REC-144@" & " " & FormatDateTime(ExpiryDate, DateFormat.ShortDate).ToString & " " & "^REC-145@" & " "
                            Exit Function
                        End If
                    Else
                        Items.Tables(0).Rows(i)("ProductionDate") = DBNull.Value
                        Items.Tables(0).Rows(i)("ExpDate") = DBNull.Value
                    End If
                    Items.Tables(0).Rows(i).AcceptChanges()

                    'Generate RTLot for IR / ASN which multiple receipt has only one RT Number      -- 7/20/2017
                    Dim myLotNumber As String
                    If Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                        myLotNumber = Format(Header.PostDate, "yyMMdd")
                    Else
                        myLotNumber = Format(Items.Tables(0).Rows(i)("ExpDate"), "yyMMdd")
                    End If

                    'Header.Type = 9:  CLID creation based on Oracle Receipt, no need to send data to Oracle
                    If Header.Type = 9 Then
                        'Auto Generate Lot Number for IR Receipt if no Lot found from Source org                         '12/18/2017
                        If Header.OrderType = 4 AndAlso Items.Tables(0).Rows(i)("LotControl") = "Y" Then
                            If FixNull(Items.Tables(0).Rows(i)("RTLot")) = "" Then
                                If Not Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                                    myLotNumber = CLng(Format(Items.Tables(0).Rows(i)("ExpDate"), "yyyyMMdd"))
                                Else
                                    myLotNumber = CLng(Format(DateTime.Now, "yyyyMMdd"))
                                End If
                                Dim LotNumber As String = myLotNumber.ToString             'CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))
                                If Header.POCurrency = "CNY" Then LotNumber = LotNumber & "CN"
                                Items.Tables(0).Rows(i)("RTLot") = LotNumber
                            End If
                        End If

                        If Header.OrderType = 6 AndAlso Items.Tables(0).Rows(i)("LotControl") = "Y" Then              '03/27/2015
                            Dim LotNumber As String = myLotNumber.ToString             'CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))
                            If Header.POCurrency = "CNY" Then LotNumber = LotNumber & "CN"
                            Items.Tables(0).Rows(i)("RTLot") = LotNumber
                        End If

                        If Header.OrderType = 3 OrElse Header.OrderType = 4 Then
                            'Set EJitFlag to see if need to extract EJitID from Oracle
                            If EJitRTFlag = False Then EJitRTFlag = True

                            Dim myDR As DataRow
                            myDR = dtEJit.NewRow
                            myDR("RTNo") = Header.GRNo
                            myDR("TranID") = Items.Tables(0).Rows(i)("POLineID").ToString
                            myDR("RecType") = "PO"
                            If Header.OrderType = 4 Then myDR("RecType") = "IR"
                            dtEJit.Rows.Add(myDR)
                        End If

                    ElseIf Header.Type <> 9 Then
                        DR = Nothing
                        DR = PODetails.Tables("po_line").Select(" POLineID = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                        If DR.Length > 0 Then
                            DR(0)("USER_ID") = LoginData.UserID
                            DR(0)("DELIVERY_TYPE") = Items.Tables(0).Rows(i)("DeliveryType").ToString     'Send DeliveryType to Oracle if there has
                            DR(0)("REC_QTY") = Items.Tables(0).Rows(i)("EntryQty") / Items.Tables(0).Rows(i)("BaseUOMConv")
                            DR(0)("COMMENTS") = Items.Tables(0).Rows(i)("Operator").ToString & " " & Items.Tables(0).Rows(i)("StoragePosition").ToString & " " & Items.Tables(0).Rows(i)("BoxDesc").ToString & " " & Items.Tables(0).Rows(i)("ItemText").ToString
                            DR(0)("header_interface_id") = HeaderInterfaceID
                            DR(0)("AML_STATUS") = DBNull.Value

                            'Read COO data from CLID detail info
                            DRLine = Nothing
                            'DRLine = Items.Tables(1).Select(" PONO = '" & Items.Tables(0).Rows(i)("PONo") & "' and POItem = '" & Items.Tables(0).Rows(i)("POItem") & "'")
                            DRLine = Items.Tables(1).Select(" POLineID = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                            If DRLine.Length > 0 Then
                                DR(0)("COO") = DRLine(0)("COO").ToString
                            End If

                            If Header.Type = 1 Or Header.Type = 6 Then       'Normal PO / ASN
                                'Send AML_Status to Oracle interface   -----   AML_Status = 1: AML Match; AML_Status = 2: AML Not Match ;
                                ' --- AML_Status = 1 or 2 (MSL Match for 1 & 2);  Add below two values for AML_Status  -- 01/13/2015
                                ' --- AML_Status = 3: AML Match & MSL not Match; AML_Status = 4: AML Not Match & MSL not Match   
                                If Items.Tables(0).Rows(i)("QMLStatus").ToString.ToUpper = "USERINPUT" Then
                                    DR(0)("AML_STATUS") = 2                                 'AML not Match & MSL Match
                                    If Items.Tables(0).Rows(i)("MSL") <> Items.Tables(0).Rows(i)("ReasonCode") Then
                                        DR(0)("AML_STATUS") = 4                             'AML Not Match & MSL not Match 
                                    End If
                                Else
                                    DR(0)("AML_STATUS") = 1                                  'AML Match & MSL Match
                                    If Items.Tables(0).Rows(i)("MSL") <> Items.Tables(0).Rows(i)("ReasonCode") Then
                                        DR(0)("AML_STATUS") = 3                             'AML Match & MSL not Match
                                    End If
                                End If
                                Items.Tables(0).Rows(i)("MSL") = Items.Tables(0).Rows(i)("ReasonCode")   'Read MSL from user input field
                                Items.Tables(0).Rows(i)("ReasonCode") = ""

                                'Auto Generate Lot Number for ASN PO Receipt 
                                If Header.Type = 6 AndAlso Items.Tables(0).Rows(i)("LotControl") = "Y" Then              '10/25/2013
                                    Dim LotNumber As String = myLotNumber.ToString             'CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))
                                    If Header.POCurrency = "CNY" Then LotNumber = LotNumber & "CN"
                                    Items.Tables(0).Rows(i)("RTLot") = LotNumber
                                End If

                            ElseIf Header.Type = 2 Then   'OSP PO
                                ' Change SubInventory/Locator if necessary
                                If DR(0)("WIP_COMPLETION_SUBINVENTORY").ToString = Items.Tables(0).Rows(i)("SubInventory").ToString AndAlso DR(0)("WIP_COMPLETION_LOCATOR").ToString = Items.Tables(0).Rows(i)("Locator").ToString Then
                                    DR(0)("WIP_COMPLETION_SUBINVENTORY") = ""
                                    DR(0)("WIP_COMPLETION_LOCATOR") = ""
                                Else
                                    DR(0)("WIP_COMPLETION_SUBINVENTORY") = Items.Tables(0).Rows(i)("SubInventory").ToString
                                    DR(0)("WIP_COMPLETION_LOCATOR") = Items.Tables(0).Rows(i)("Locator").ToString
                                End If
                            ElseIf Header.Type = 3 Or Header.Type = 7 Then   ' 3: Kanban/ejit PO; 7: ASN for ejit PO; 
                                DR(0)("REC_PRIMARY_QTY") = Items.Tables(0).Rows(i)("EntryQty")
                                DR(0)("lot_expiration_date") = Items.Tables(0).Rows(i)("ExpDate")

                                ' Change Locator if necessary
                                If FixNull(DR(0)("DESTINATION_LOCATOR")) <> Items.Tables(0).Rows(i)("Locator").ToString Then
                                    DR(0)("LOCATOR_ID") = 0
                                    DR(0)("DESTINATION_LOCATOR") = Items.Tables(0).Rows(i)("Locator").ToString
                                End If

                            ElseIf Header.Type = 4 Then   ' 4: IR / ISO
                                DR(0)("REC_PRIMARY_QTY") = Items.Tables(0).Rows(i)("EntryQty")
                                DR(0)("lot_expiration_date") = Items.Tables(0).Rows(i)("ExpDate")

                                'Auto Generate Lot Number for IR Receipt if no Lot found from Source org
                                If Items.Tables(0).Rows(i)("LotControl") = "Y" AndAlso FixNull(DR(0)("rt_lot_number")) = "" Then      '10/25/2013
                                    If Not Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                                        myLotNumber = CLng(Format(Items.Tables(0).Rows(i)("ExpDate"), "yyyyMMdd"))
                                    Else
                                        myLotNumber = CLng(Format(DateTime.Now, "yyyyMMdd"))
                                    End If
                                    Dim LotNumber As String = myLotNumber.ToString             'CLng(Format(DateTime.Now, "yyyyMMddHHmmss"))
                                    If Header.POCurrency = "CNY" Then LotNumber = LotNumber & "CN"

                                    If DR(0)("ir_type") <> "NO_IR" Then DR(0)("rt_lot_number") = LotNumber
                                    Items.Tables(0).Rows(i)("RTLot") = LotNumber
                                End If

                            ElseIf Header.Type = 5 Then   ' 5: Indirect PO 
                                DR(0)("REC_PRIMARY_QTY") = Items.Tables(0).Rows(i)("EntryQty")
                            End If

                            DR(0).AcceptChanges()
                            DR(0).SetAdded()
                            dsPOData.Tables("po_line").ImportRow(DR(0))
                        End If
                    End If
                End If
            Next

            'Get EJit ID for EJit PO / IR if Receipt Type = 9               04/22/2014
            If Header.Type = 9 AndAlso EJitRTFlag = True Then
                Try
                    GetEJitData(LoginData, dsEJit)
                Catch ex As Exception
                    ErrorLogging("Receiving-Call-GetEJitData", LoginData.User, "RTNo: " & GRNo & "; " & ex.Message & ex.Source, "E")
                End Try
            End If


            'Header.Type = 9:  CLID creation based on Oracle Receipt, no need to send data to Oracle
            If Header.Type <> 9 Then
                For i = 0 To dsPOData.Tables("po_line").Rows.Count - 1
                    If dsPOData.Tables("po_line").Rows(i).RowState = DataRowState.Unchanged Then
                        dsPOData.Tables("po_line").Rows(i).SetAdded()
                    End If
                Next

                Try
                    'Make sure only 1 row in table "po_header"
                    Dim l_row As Integer
                    If dsPOData.Tables("po_header").Rows.Count > 1 Then
                        For l_row = dsPOData.Tables("po_header").Rows.Count - 1 To 1 Step -1
                            dsPOData.Tables("po_header").Rows.RemoveAt(l_row)
                        Next
                    End If
                    If dsPOData.Tables("po_line").Rows.Count > 0 Then
                        If Header.Type = 2 Then
                            PostGR = Post_OSP_PO(LoginData, dsPOData, GroupID, Header)
                        Else
                            PostGR = Post_Receive(LoginData, dsPOData, GroupID, Header)
                        End If

                    End If
                Catch ex As Exception
                    ErrorLogging("Receiving-PostGR-OraPKG", LoginData.User.ToUpper, "BatchID: " & GroupID & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                    Return Nothing
                End Try
                If PostGR.GRStatus Is Nothing OrElse PostGR.GRStatus = "E" Then Exit Function

                GRNo = PostGR.GRNo
                PostDate = PostGR.PostDate
                If PostGR.GRMsg Is Nothing OrElse PostGR.GRMsg.Tables.Count = 0 Then
                ElseIf PostGR.GRMsg.Tables.Count > 0 Then
                    GRResult.Merge(PostGR.GRMsg.Tables(0))
                    PostGR.GRMsg = Nothing
                End If

                'Check if there has OSP DJs which have OQA operation before DJ completion or not, if Yes, then don't need to create record in eTrace
                If Header.Type = 2 Then
                    Dim OQAFlag As String = "N"   'N: No OQA operation;   Y: Have OQA operation 
                    OQAFlag = dsPOData.Tables("po_line").Rows(0)("oqa_flag").ToString
                    If OQAFlag = "Y" Then
                        PostGR.CLIDs = New DataSet
                        PostGR.CLIDs = Items
                        Exit Function
                    End If
                End If

                'For Indirect PO Receipt, no need to create record in eTrace    11/15/2013
                If Header.Type = 5 Then
                    PostGR.CLIDs = New DataSet
                    PostGR.CLIDs = Items
                    Exit Function
                End If
            End If


            Dim myCommand As SqlClient.SqlCommand
            Dim CLMasterSQLCommand As SqlClient.SqlCommand
            Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

            Dim ra, k As Integer
            Dim NextCLID As String
            Dim SourceCLID As String = ""
            Dim myStatusCode As String = "1"

            CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_CLMaster ( CLID,OrgCode,StatusCode,MaterialNo,MaterialRevision,Qty,UOM,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RecDocNo,RecDocItem,RTLot,CreatedBy,CreatedOn,VendorID,RecDate,Printed,ProcessID,ExpDate,ProdDate,RoHS,PurOrdNo,PurOrdItem,DeliveryType,InvoiceNo,BillofLading,DN,HeaderText,ReasonCode, StockType, MatSuffix1,MatSuffix2,MatSuffix3,MaterialDesc,VendorName,VendorPN,StoragePosition,SLOC,StorageBin,PredefinedSubInv,PredefinedLocator,Operator,IsTraceable,AddlText,MatDocNo,MatDocItem,Manufacturer,ManufacturerPN,QMLStatus, AddlData, Stemp, MSL, BoxID, SourceID, LastTransaction, SourceCLID, MCPosition ) values (@CLID, @OrgCode, @StatusCode, @MaterialNo,@MaterialRevision,@Qty,@UOM,@QtyBaseUOM,@BaseUOM,@DateCode,@LotNo,@COO,@RecDocNo,@RecDocItem,@RTLot,@CreatedBy,getdate(),@VendorID,@RecDate,@Printed,@ProcessID,@ExpDate,@ProdDate,@RoHS,@PurOrdNo,@PurOrdItem,@DeliveryType,@InvoiceNo,@BillofLading,@DN,@HeaderText,@ReasonCode,@StockType, @MatSuffix1,@MatSuffix2,@MatSuffix3,@MaterialDesc,@VendorName,@VendorPN,@StoragePosition,@SLOC,@StorageBin,@PredefinedSubInv,@PredefinedLocator,@Operator,@IsTraceable,@AddlText,@MatDocNo,@MatDocItem,@Manufacturer,@ManufacturerPN,@QMLStatus,@AddlData,@Stemp,@MSL,@BoxID,@SourceID,@LastTransaction, @SourceCLID, @MCPosition ) ", myConn)
            CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
            CLMasterSQLCommand.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10, "StatusCode")                     'Add StatusCode as parameter for CLID Verification
            CLMasterSQLCommand.Parameters.Add("@MaterialNo", SqlDbType.VarChar, 30, "MaterialNo")
            CLMasterSQLCommand.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 20, "MaterialRevision")
            CLMasterSQLCommand.Parameters.Add("@Qty", SqlDbType.Decimal, 25, "Qty")
            CLMasterSQLCommand.Parameters.Add("@UOM", SqlDbType.VarChar, 10, "UOM")
            CLMasterSQLCommand.Parameters.Add("@QtyBaseUOM", SqlDbType.Decimal, 25, "QtyBaseUOM")
            CLMasterSQLCommand.Parameters.Add("@BaseUOM", SqlDbType.VarChar, 10, "BaseUOM")
            CLMasterSQLCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 20, "DateCode")
            CLMasterSQLCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 50, "LotNo")
            CLMasterSQLCommand.Parameters.Add("@COO", SqlDbType.VarChar, 20, "COO")
            CLMasterSQLCommand.Parameters.Add("@RecDocNo", SqlDbType.VarChar, 20, "RecDocNo")
            CLMasterSQLCommand.Parameters.Add("@RecDocItem", SqlDbType.VarChar, 10, "RecDocItem")
            CLMasterSQLCommand.Parameters.Add("@RTLot", SqlDbType.VarChar, 50, "RTLot")
            CLMasterSQLCommand.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
            CLMasterSQLCommand.Parameters.Add("@Printed", SqlDbType.VarChar, 100, "Printed")
            CLMasterSQLCommand.Parameters.Add("@VendorID", SqlDbType.VarChar, 50, "VendorID")
            CLMasterSQLCommand.Parameters.Add("@RecDate", SqlDbType.DateTime, 8, "RecDate")
            CLMasterSQLCommand.Parameters.Add("@ProcessID", SqlDbType.Int, 4, "ProcessID")
            CLMasterSQLCommand.Parameters.Add("@RoHS", SqlDbType.VarChar, 20, "RoHS")
            CLMasterSQLCommand.Parameters.Add("@PurOrdNo", SqlDbType.VarChar, 20, "PurOrdNo")
            CLMasterSQLCommand.Parameters.Add("@PurOrdItem", SqlDbType.VarChar, 20, "PurOrdItem")
            CLMasterSQLCommand.Parameters.Add("@DeliveryType", SqlDbType.VarChar, 20, "DeliveryType")
            CLMasterSQLCommand.Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 25, "InvoiceNo")
            CLMasterSQLCommand.Parameters.Add("@BillofLading", SqlDbType.VarChar, 25, "BillofLading")
            CLMasterSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20, "OrgCode")
            CLMasterSQLCommand.Parameters.Add("@DN", SqlDbType.VarChar, 25, "DN")
            CLMasterSQLCommand.Parameters.Add("@HeaderText", SqlDbType.VarChar, 25, "HeaderText")
            CLMasterSQLCommand.Parameters.Add("@ExpDate", SqlDbType.SmallDateTime, 4, "ExpDate")
            CLMasterSQLCommand.Parameters.Add("@ProdDate", SqlDbType.SmallDateTime, 4, "ProdDate")
            CLMasterSQLCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 20, "ReasonCode")
            CLMasterSQLCommand.Parameters.Add("@StockType", SqlDbType.VarChar, 10, "StockType")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix1", SqlDbType.VarChar, 10, "MatSuffix1")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix2", SqlDbType.VarChar, 10, "MatSuffix2")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix3", SqlDbType.VarChar, 10, "MatSuffix3")
            CLMasterSQLCommand.Parameters.Add("@MaterialDesc", SqlDbType.VarChar, 250, "MaterialDesc")
            CLMasterSQLCommand.Parameters.Add("@VendorName", SqlDbType.VarChar, 100, "VendorName")
            CLMasterSQLCommand.Parameters.Add("@VendorPN", SqlDbType.VarChar, 100, "VendorPN")
            CLMasterSQLCommand.Parameters.Add("@StoragePosition", SqlDbType.VarChar, 50, "StoragePosition")
            CLMasterSQLCommand.Parameters.Add("@SLOC", SqlDbType.VarChar, 20, "SLOC")                             'SLOC, StorageBin
            CLMasterSQLCommand.Parameters.Add("@StorageBin", SqlDbType.VarChar, 20, "StorageBin")
            CLMasterSQLCommand.Parameters.Add("@PredefinedSubInv", SqlDbType.VarChar, 20, "PredefinedSubInv")
            CLMasterSQLCommand.Parameters.Add("@PredefinedLocator", SqlDbType.VarChar, 20, "PredefinedLocator")
            CLMasterSQLCommand.Parameters.Add("@Operator", SqlDbType.VarChar, 50, "Operator")
            CLMasterSQLCommand.Parameters.Add("@IsTraceable", SqlDbType.VarChar, 10, "IsTraceable")
            CLMasterSQLCommand.Parameters.Add("@AddlText", SqlDbType.VarChar, 50, "AddlText")
            CLMasterSQLCommand.Parameters.Add("@MatDocNo", SqlDbType.VarChar, 20, "MatDocNo")
            CLMasterSQLCommand.Parameters.Add("@MatDocItem", SqlDbType.VarChar, 10, "MatDocItem")
            CLMasterSQLCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 100, "Manufacturer")
            CLMasterSQLCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 100, "ManufacturerPN")
            CLMasterSQLCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
            CLMasterSQLCommand.Parameters.Add("@AddlData", SqlDbType.VarChar, 20, "AddlData")
            CLMasterSQLCommand.Parameters.Add("@Stemp", SqlDbType.VarChar, 50, "Stemp")
            CLMasterSQLCommand.Parameters.Add("@MSL", SqlDbType.VarChar, 50, "MSL")
            CLMasterSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
            CLMasterSQLCommand.Parameters.Add("@SourceID", SqlDbType.VarChar, 50, "SourceID")
            CLMasterSQLCommand.Parameters.Add("@LastTransaction", SqlDbType.VarChar, 100, "LastTransaction")
            CLMasterSQLCommand.Parameters.Add("@SourceCLID", SqlDbType.VarChar, 20, "SourceCLID")
            CLMasterSQLCommand.Parameters.Add("@MCPosition", SqlDbType.VarChar, 50, "MCPosition")
            CLMasterSQLCommand.Parameters("@LastTransaction").Value = "Receiving"
            CLMasterSQLCommand.Parameters("@SourceCLID").Value = DBNull.Value
            CLMasterSQLCommand.Parameters("@MCPosition").Value = DBNull.Value

            'Call Stored Procedure to get NextID
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            If Header.OrderType = 2 Then
                myCommand.Parameters.AddWithValue("@TypeID", "PCBID")
            Else
                myCommand.Parameters.AddWithValue("@TypeID", "CLID")
            End If
            myCommand.CommandTimeout = TimeOut_M5


            Dim RecQty As Decimal
            Dim DocItem As Integer = 0
            Dim ItemIndex As Integer = 0
            Dim NoOfPackage As Integer = 0

            Dim lblPrint As String
            If PrintLabels = True Then  '!' To ensure timely feedback, Do label printing later
                lblPrint = "True"
            Else
                lblPrint = "Disabled"
            End If


            Try
                myConn.Open()

                'Read Receiving DashBoard Flag from Table T_Config to decide if need to Save PalletID info
                SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC001'")
                Sqlcmd = New SqlClient.SqlCommand(SqlStr, myConn)
                RecConfig = FixNull(Sqlcmd.ExecuteScalar)


                'For Normal PO, check if the Vendor need to do CLID AML Verification, if Yes, set StatusCode = "9" insteads of "1"            -- 01/18/2017
                Dim tmpVendor As String = "", CLIDVerifyFlag As Boolean = False
                If Header.OrderType = 1 Or Header.OrderType = 6 Then
                    Dim POVendor As String = PODetails.Tables("po_header").Rows(0)("VENDOR_NUMBER").ToString
                    SqlStr = String.Format("Select ProcessID as VendorNo from T_SysLOV with (nolock) where Name = 'Vendors for CLID Verification' and ProcessID = '" & POVendor & "'")
                    Sqlcmd = New SqlClient.SqlCommand(SqlStr, myConn)
                    tmpVendor = FixNull(Sqlcmd.ExecuteScalar)
                    'If tmpVendor <> "" Then
                    '    myStatusCode = "9"
                    '    PostGR.GRStatus = "Y/V"                      'Set Flag to record the RT which need to do AML Verification
                    'End If
                End If

               '2019-02-16 Added Sample Location Column and Printed
                Dim dsSampleLoc As DataSet = New DataSet
                Dim dtSampleLoc As DataTable = New DataTable
                dtSampleLoc.Columns.Add(New Data.DataColumn("OrgCode", System.Type.GetType("System.String")))
                dtSampleLoc.Columns.Add(New Data.DataColumn("Item", System.Type.GetType("System.String")))
                dtSampleLoc.Columns.Add(New Data.DataColumn("SampleLocation", System.Type.GetType("System.String")))
                dsSampleLoc.Tables.Add(dtSampleLoc)
                Dim drSampleloc() As DataRow

                For i = 0 To Items.Tables(0).Rows.Count - 1
                    If Items.Tables(0).Rows(i)("EntryQty") Is DBNull.Value OrElse Items.Tables(0).Rows(i)("EntryQty") = "" Then
                        Items.Tables(0).Rows(i)("EntryQty") = 0

                    ElseIf Convert.ToDecimal(Items.Tables(0).Rows(i)("EntryQty")) > 0 Then
                        ItemIndex = ItemIndex + 1

                        'For Normal PO Receipt, RTLot=RTNo, for KB_PO / EJ_PO / KB_IR / EJ_IR which receipt routing = Direct Delivery, get the RTLot from Package
                        Dim RTLot As String = ""
                        Dim MCPosition As String = ""
                        Dim TransactionID As String = Items.Tables(0).Rows(i)("POLineID")

                        DRLine = Nothing
                        DRLine = PODetails.Tables("po_line").Select(" POLineID = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                        If DRLine.Length > 0 Then

                            Dim Flag As Boolean = False
                            Dim msgDR() As DataRow = Nothing
                            'If Header.Type = 2 OrElse Header.Type = 9 Then   'For OSP PO, no need to record Transaction ID in eTrace
                            If Header.Type = 9 Then   'For OSP PO, also need to record Transaction ID in eTrace, this is to avoid duplicated record generated for Receipt Type 9 with OSP   -- 06/16/2014
                                Flag = True
                            ElseIf Header.Type <> 9 Then
                                msgDR = GRResult.Tables(0).Select(" msg_type = 'TRANSACTION_ID' and msg_id = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                                If msgDR.Length > 0 Then Flag = True
                            End If

                            Dim SIPFlag, RoHSFlag, IQCFlag As String
                            Dim PORoutingID As Integer
                            Dim myEJITID As String = ""
                            Dim drEJit() As DataRow = Nothing
                            Dim DateCode, LotNo, VarDateCode, VarLotNo, InspectionRT As String, OldInspectionRT As String

                            If Flag = True Then
                                If Header.OrderType <> 2 Then
                                    'Get MCPosition from item PIM setting for Non-OSP PO Receipt
                                    MCPosition = DRLine(0)("item_pim_sp").ToString

                                    If Header.Type = 9 Then
                                        DocItem = DRLine(0)("transaction_id").ToString         'TransactionID recorded in field POLineID
                                        RTLot = DRLine(0)("rt_lot_number").ToString
                                        If Header.OrderType = 4 Then
                                            'For Normal IR, get Lot from Source Org, for KB/Ejit IR, get Lot from Oracle result
                                            If RTLot = "" Then RTLot = Items.Tables(0).Rows(i)("RTLot").ToString
                                        ElseIf Header.OrderType = 6 Then
                                            RTLot = GRNo & "-" & Items.Tables(0).Rows(i)("RTLot").ToString
                                        End If

                                        PORoutingID = DRLine(0)("receiving_routing_id")
                                        SIPFlag = DRLine(0)("sip_flag").ToString.ToUpper
                                        RoHSFlag = DRLine(0)("rohs_flag").ToString.ToUpper
                                        IQCFlag = DRLine(0)("iqc_flag").ToString.ToUpper           'Add IQC Flag 

                                        'Record EJit ID for EJit RT              04/22/2014
                                        If EJitRTFlag = True Then
                                            If dsEJit.Tables.Count > 0 AndAlso dsEJit.Tables(0).Rows.Count > 0 Then
                                                drEJit = dsEJit.Tables(0).Select(" TranID = '" & TransactionID & "'")
                                                If drEJit.Length > 0 AndAlso drEJit(0)("eJITID").ToString <> "" Then
                                                    myEJITID = drEJit(0)("eJITID").ToString
                                                    If RTLot = "" AndAlso Items.Tables(0).Rows(i)("LotControl") = "Y" Then
                                                        RTLot = drEJit(0)("RTLot").ToString
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        DocItem = msgDR(0)("text").ToString
                                        RTLot = msgDR(0)("rt_lot_number").ToString
                                        If Header.Type = 4 Then
                                            'For Normal IR, get Lot from Source Org, for KB/Ejit IR, get Lot from Oracle result
                                            If DRLine(0)("ir_type") = "NO_IR" Then
                                                RTLot = DRLine(0)("rt_lot_number").ToString
                                                If RTLot = "" Then RTLot = Items.Tables(0).Rows(i)("RTLot").ToString
                                            End If
                                        ElseIf Header.Type = 6 Then
                                            RTLot = GRNo & "-" & Items.Tables(0).Rows(i)("RTLot").ToString
                                        End If

                                        PORoutingID = msgDR(0)("rt_routing_id")
                                        SIPFlag = msgDR(0)("sip_flag").ToString.ToUpper
                                        RoHSFlag = msgDR(0)("rohs_flag").ToString.ToUpper
                                        IQCFlag = msgDR(0)("iqc_flag").ToString.ToUpper           'Add IQC Flag 

                                        If (Header.Type = 3 AndAlso DRLine(0)("po_type") = "KB_PO") OrElse (Header.Type = 7 AndAlso DRLine(0)("po_type") = "KB_ASNPO") _
                                         OrElse (Header.Type = 4 AndAlso DRLine(0)("ir_type") = "KB_IR") Then
                                            DocItem = CInt(Items.Tables(0).Rows(i)("ReasonCode"))   'ItemIndex
                                        End If
                                    End If

                                    'Generate Oracle Lot in eTrace if there is blank   10/17/2014
                                    If RTLot = "" AndAlso Items.Tables(0).Rows(i)("LotControl") = "Y" Then
                                        RTLot = GRNo
                                        If Header.POCurrency = "CNY" Then RTLot = RTLot & "CN"
                                    End If

                                    'Get RT Routing from Oracle after RT generated and save it to eTrace
                                    If PORoutingID = 2 Then
                                        Items.Tables(0).Rows(i)("StockType") = IQCFlag          'Default value from Oracle RT
                                        If IQCFlag = "" Then
                                            If SIPFlag = "Q" Then
                                                Items.Tables(0).Rows(i)("StockType") = "Q"
                                                If RoHSFlag = "Y" Then Items.Tables(0).Rows(i)("StockType") = "QT"

                                                If Items.Tables(0).Rows(i)("QMLStatus").ToString.ToUpper = "USERINPUT" Then
                                                    Items.Tables(0).Rows(i)("StockType") = "QV"
                                                    If RoHSFlag = "Y" Then Items.Tables(0).Rows(i)("StockType") = "QT/QV"
                                                End If
                                            ElseIf SIPFlag = "S" Then
                                                Items.Tables(0).Rows(i)("StockType") = "S"
                                                'For SIP Inspection, don't need to consider RoHS Test Flag      --- 10/10/2013
                                                'If RoHSFlag = "Y" Then Items.Tables(0).Rows(i)("StockType") = "ST"

                                                If Items.Tables(0).Rows(i)("QMLStatus").ToString.ToUpper = "USERINPUT" Then
                                                    Items.Tables(0).Rows(i)("StockType") = "SV"
                                                    'For SIP Inspection, don't need to consider RoHS Test Flag      --- 10/10/2013
                                                    'If RoHSFlag = "Y" Then Items.Tables(0).Rows(i)("StockType") = "ST/SV"
                                                End If
                                            End If
                                        End If
                                    Else
                                        Items.Tables(0).Rows(i)("StockType") = "FTS"
                                    End If
                                End If

                                Dim tmpStr As String = ""
                                If Header.Type = 9 Then
                                    tmpStr = " PONO = '" & Items.Tables(0).Rows(i)("PONo") & "' and POItem = '" & Items.Tables(0).Rows(i)("POItem") & "' and POLineID = '" & TransactionID & "'"
                                Else
                                    'Record Transaction ID in eTrace for all PO Types, included OSP PO -- 06/16/2014
                                    TransactionID = msgDR(0)("text").ToString
                                    tmpStr = " PONO = '" & Items.Tables(0).Rows(i)("PONo") & "' and POItem = '" & Items.Tables(0).Rows(i)("POItem") & "'"
                                End If
                                DR = Nothing
                                DR = Items.Tables(1).Select(tmpStr)
                                If DR.Length > 0 Then
                                    NoOfPackage = 0                                   'Initial NoOfPackage to zero before loop
                                    DateCode = ""
                                    LotNo = ""
                                    VarDateCode = ""
                                    VarLotNo = ""
                                    InspectionRT = ""
                                    OldInspectionRT = ""
                                    For j = 0 To DR.Length - 1
                                        If IsNumeric(DR(j)("Qty")) = False Then DR(j)("Qty") = 0
                                        If DR(j)("Qty") <= 0 Then
                                            DR(j).Delete()
                                        Else

                                            NextCLID = ""
                                            NoOfPackage = NoOfPackage + 1             'For each New CLID, add one to NoOfPackage

                                            'Vendor CLID is provided from excel upload file    -- 4/25/2016
                                            If (Header.OrderType = 1 Or Header.OrderType = 3 Or Header.OrderType = 6 Or Header.OrderType = 7) _
                                                 AndAlso Header.CLIDFlag = True AndAlso DR(j)("LID").ToString <> "" Then
                                                myStatusCode = "1"
                                                NextCLID = DR(j)("LID").ToString
                                            ElseIf Header.OrderType = 2 AndAlso DR(j)("LID").ToString <> "" Then   'For OSP DJ, CartonID/PalletID stands for CLID/BoxID
                                                myStatusCode = "1"
                                                NextCLID = DR(j)("LID").ToString
                                                DR(j)("LotNo") = ""
                                                DR(j)("RefCLID") = ""
                                                DR(j)("Status") = "X"    'Set Flag for BoxID/PalletID then not to print labels
                                            ElseIf Header.OrderType = 4 AndAlso DR(j)("LID").ToString <> "" AndAlso PrintLabels = False Then      'Use Original CLID for IR/ISO if users do not print labels
                                                myStatusCode = "1"
                                                NextCLID = DR(j)("LID").ToString
                                            Else
                                                'Try up to 5 times when failed getting next id
                                                k = 0
                                                While (k < 5 And NextCLID = "")
                                                    Try
                                                        ra = myCommand.ExecuteNonQuery()
                                                        NextCLID = myCommand.Parameters(0).Value
                                                    Catch ex As Exception
                                                        k = k + 1
                                                        ErrorLogging("Receiving-PostGR-GetNextID", "Deadlocked? " & Str(k), "Failed getting next ID; RTNo: " & GRNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                                                    End Try
                                                End While
                                                If NextCLID = "" Then
                                                    ErrorLogging("Receiving-PostGR-GetNextID", LoginData.User.ToUpper, Str(k) & " Failed getting next ID; RTNo: " & GRNo & " with Receipt type " & Header.Type, "I")
                                                    Exit For
                                                End If
                                                If Header.OrderType = 1 Or Header.OrderType = 6 Then
                                                    If tmpVendor <> "" Then
                                                        myStatusCode = "9"
                                                        CLIDVerifyFlag = True
                                                    End If
                                                End If
                                            End If

                                            'Clear SourceCLID to make sure the data correct
                                            SourceCLID = ""
                                            If Header.OrderType = 4 Then
                                                If DR(j)("LID").ToString <> "" OrElse DR(j)("RefCLID").ToString <> "" Then      'Record CLID for IR/ISO 
                                                    Dim LID As String
                                                    If DR(j)("LID").ToString <> "" Then
                                                        LID = DR(j)("LID").ToString
                                                    Else
                                                        LID = DR(j)("RefCLID").ToString
                                                    End If
                                                    IRDataRow = IRLabels.Tables("CLIDs").NewRow()
                                                    IRDataRow("CLID") = LID
                                                    Try
                                                        IRLabels.Tables("CLIDs").Rows.Add(IRDataRow)
                                                    Catch ex As Exception
                                                    End Try

                                                    'Save the Source CLID and record into T_CLMaster later   "Yudy20120402
                                                    SourceCLID = LID
                                                End If
                                            End If

                                            RecQty = DR(j)("Qty")
                                            If Items.Tables(0).Rows(i)("BaseUOM") <> Items.Tables(0).Rows(i)("UOM") Then
                                                RecQty = RecQty / Items.Tables(0).Rows(i)("BaseUOMConv")
                                            End If

                                            'If Header.Type = 2 OrElse (Header.Type = 9 AndAlso Line_type = 3) Then        'OSP DJ
                                            If Header.OrderType = 2 Then        'OSP DJ
                                                DocItem = ItemIndex
                                                CLMasterSQLCommand.Parameters("@PurOrdNo").Value = Items.Tables(0).Rows(i)("ReasonCode")       'OSP DJ
                                                CLMasterSQLCommand.Parameters("@PurOrdItem").Value = DBNull.Value
                                                CLMasterSQLCommand.Parameters("@MatDocNo").Value = Items.Tables(0).Rows(i)("PONo")
                                                CLMasterSQLCommand.Parameters("@MatDocItem").Value = Items.Tables(0).Rows(i)("POItem")
                                                CLMasterSQLCommand.Parameters("@ReasonCode").Value = ""

                                                'Add SubInventory/Locator for OSP Assembly      Yudy 11/03/2009
                                                CLMasterSQLCommand.Parameters("@SLOC").Value = Items.Tables(0).Rows(i)("SubInventory").ToString
                                                CLMasterSQLCommand.Parameters("@StorageBin").Value = Items.Tables(0).Rows(i)("Locator").ToString
                                                CLMasterSQLCommand.Parameters("@PredefinedSubInv").Value = DBNull.Value
                                                CLMasterSQLCommand.Parameters("@PredefinedLocator").Value = DBNull.Value
                                                If DR(j)("BoxID").ToString <> "" Then
                                                    CLMasterSQLCommand.Parameters("@BoxID").Value = DR(j)("BoxID").ToString
                                                Else
                                                    CLMasterSQLCommand.Parameters("@BoxID").Value = DBNull.Value
                                                End If

                                                'For SubAssembly, Record WIP Revision and save it to LotNo field if T_Config set value = YES
                                                CLMasterSQLCommand.Parameters("@LotNo").Value = ""
                                                If Items.Tables(0).Rows(i)("RevControl") = "N" Then   'No Revision control for SubAssembly
                                                    Dim RevConfig As String
                                                    SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'CLID008'")
                                                    Sqlcmd = New SqlClient.SqlCommand(SqlStr, myConn)
                                                    RevConfig = FixNull(Sqlcmd.ExecuteScalar)
                                                    If RevConfig = "YES" Then
                                                        'Read DJ BOM revision first, if not found for Rework DJ, then read Item Master Rev
                                                        Dim DJRev As String = DRLine(0)("WIP_REVISION").ToString
                                                        If DJRev = "" Then DJRev = DRLine(0)("item_rev").ToString
                                                        CLMasterSQLCommand.Parameters("@LotNo").Value = DJRev
                                                    End If
                                                End If

                                            Else

                                                'Combine DateCode & LotNo, and print it out on RT Slip which requested by PH  -- 05/16/2014
                                                Dim tmpDC As String = DR(j)("DateCode").ToString.Trim
                                                Dim tmpLN As String = DR(j)("LotNo").ToString.Trim
                                                If tmpDC <> "" Then
                                                    If DateCode = "" Then
                                                        DateCode = tmpDC
                                                    Else
                                                        If Not DateCode.Contains(tmpDC) Then
                                                            DateCode = DateCode & ";" & tmpDC
                                                        End If
                                                    End If
                                                End If
                                                If tmpLN <> "" Then
                                                    If LotNo = "" Then
                                                        LotNo = tmpLN
                                                    Else
                                                        If Not LotNo.Contains(tmpLN) Then
                                                            LotNo = LotNo & ";" & tmpLN
                                                        End If
                                                    End If
                                                End If

                                                If j = 0 Then
                                                    VarDateCode = tmpDC
                                                    VarLotNo = tmpLN
                                                    If VarDateCode = "" And VarLotNo = "" Then
                                                        InspectionRT = ""
                                                        OldInspectionRT = InspectionRT
                                                    Else
                                                    InspectionRT = GetInspectionRT(LoginData.OrgCode, Items.Tables(0).Rows(i)("MaterialNo"), Items.Tables(0).Rows(i)("VendorNo"), VarDateCode, VarLotNo)
                                                    OldInspectionRT = InspectionRT
                                                    End If
                                                Else
                                                    If OldInspectionRT <> "" Then
                                                        If VarDateCode <> tmpDC Or VarLotNo <> tmpLN Then
                                                            VarDateCode = tmpDC
                                                            VarLotNo = tmpLN
                                                            If VarDateCode = "" And VarLotNo = "" Then
                                                                InspectionRT = ""
                                                                OldInspectionRT = InspectionRT
                                                            Else
                                                            InspectionRT = GetInspectionRT(LoginData.OrgCode, Items.Tables(0).Rows(i)("MaterialNo"), Items.Tables(0).Rows(i)("VendorNo"), VarDateCode, VarLotNo)
                                                            OldInspectionRT = InspectionRT
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                CLMasterSQLCommand.Parameters("@LotNo").Value = DR(j)("LotNo").ToString
                                                CLMasterSQLCommand.Parameters("@PurOrdNo").Value = Items.Tables(0).Rows(i)("PONo")         'Header.PONo
                                                CLMasterSQLCommand.Parameters("@PurOrdItem").Value = Items.Tables(0).Rows(i)("POItem")
                                                CLMasterSQLCommand.Parameters("@MatDocNo").Value = GRNo
                                                CLMasterSQLCommand.Parameters("@MatDocItem").Value = DocItem
                                                CLMasterSQLCommand.Parameters("@ReasonCode").Value = Items.Tables(0).Rows(i)("ReasonCode")
                                                CLMasterSQLCommand.Parameters("@PredefinedSubInv").Value = Items.Tables(0).Rows(i)("PredefinedSubInv")
                                                CLMasterSQLCommand.Parameters("@PredefinedLocator").Value = Items.Tables(0).Rows(i)("PredefinedLocator")
                                                CLMasterSQLCommand.Parameters("@BoxID").Value = DBNull.Value
                                                CLMasterSQLCommand.Parameters("@MCPosition").Value = MCPosition

                                                If Header.Type = 3 OrElse Header.Type = 7 OrElse (Header.Type = 4 AndAlso DRLine(0)("ir_type") <> "NO_IR") Then
                                                    'Make sure the SubInv / Locator is not blank for KB/eJit
                                                    If Items.Tables(0).Rows(i)("SubInventory").ToString = "" Then Items.Tables(0).Rows(i)("SubInventory") = DRLine(0)("DESTINATION_SUBINVENTORY").ToString
                                                    If Items.Tables(0).Rows(i)("Locator").ToString = "" Then Items.Tables(0).Rows(i)("Locator") = DRLine(0)("DESTINATION_LOCATOR").ToString

                                                    CLMasterSQLCommand.Parameters("@SLOC").Value = Items.Tables(0).Rows(i)("SubInventory").ToString
                                                    CLMasterSQLCommand.Parameters("@StorageBin").Value = Items.Tables(0).Rows(i)("Locator").ToString
                                                Else
                                                    CLMasterSQLCommand.Parameters("@SLOC").Value = DBNull.Value
                                                    CLMasterSQLCommand.Parameters("@StorageBin").Value = DBNull.Value

                                                    'Get SubInv / Locator for EJit RT              04/22/2014
                                                    If myEJITID <> "" Then
                                                        CLMasterSQLCommand.Parameters("@SLOC").Value = drEJit(0)("SubInv").ToString
                                                        CLMasterSQLCommand.Parameters("@StorageBin").Value = drEJit(0)("Locator").ToString
                                                        If Items.Tables(0).Rows(i)("LotControl") = "Y" Then
                                                            If Items.Tables(0).Rows(i)("ExpControl") = "Y" Then Items.Tables(0).Rows(i)("ExpDate") = CDate(drEJit(0)("ExpDate"))
                                                        End If
                                                    End If
                                                End If
                                            End If

                                            'For Normal PO Receipt, RTLot=RTNo, for KB_PO / EJ_PO / KB_IR / EJ_IR which receipt routing = Direct Delivery, get the RTLot from Package
                                            CLMasterSQLCommand.Parameters("@RTLot").Value = ""
                                            If Items.Tables(0).Rows(i)("LotControl") = "Y" Then CLMasterSQLCommand.Parameters("@RTLot").Value = RTLot 'GRNo

                                            CLMasterSQLCommand.Parameters("@CLID").Value = NextCLID
                                            CLMasterSQLCommand.Parameters("@StatusCode").Value = myStatusCode                ''Add StatusCode as parameter for CLID Verification  -- 01/18/2017
                                            CLMasterSQLCommand.Parameters("@OrgCode").Value = LoginData.OrgCode
                                            CLMasterSQLCommand.Parameters("@MaterialNo").Value = Items.Tables(0).Rows(i)("MaterialNo")
                                            CLMasterSQLCommand.Parameters("@MaterialRevision").Value = Items.Tables(0).Rows(i)("MaterialRevision")
                                            CLMasterSQLCommand.Parameters("@Qty").Value = RecQty
                                            CLMasterSQLCommand.Parameters("@UOM").Value = Items.Tables(0).Rows(i)("UOM")
                                            CLMasterSQLCommand.Parameters("@QtyBaseUOM").Value = DR(j)("Qty")
                                            CLMasterSQLCommand.Parameters("@BaseUOM").Value = Items.Tables(0).Rows(i)("BaseUOM")
                                            CLMasterSQLCommand.Parameters("@DateCode").Value = DR(j)("DateCode").ToString
                                            'CLMasterSQLCommand.Parameters("@LotNo").Value = DR(j)("LotNo").ToString
                                            CLMasterSQLCommand.Parameters("@COO").Value = DR(j)("COO").ToString
                                            CLMasterSQLCommand.Parameters("@RecDocNo").Value = GRNo
                                            CLMasterSQLCommand.Parameters("@RecDocItem").Value = TransactionID               'DocItem
                                            CLMasterSQLCommand.Parameters("@CreatedBy").Value = LoginData.User.ToUpper
                                            CLMasterSQLCommand.Parameters("@Printed").Value = lblPrint
                                            CLMasterSQLCommand.Parameters("@DeliveryType").Value = Items.Tables(0).Rows(i)("DeliveryType").ToString
                                            CLMasterSQLCommand.Parameters("@VendorID").Value = Items.Tables(0).Rows(i)("VendorNo")
                                            CLMasterSQLCommand.Parameters("@RecDate").Value = PostDate    'Header.PostDate
                                            CLMasterSQLCommand.Parameters("@RoHS").Value = Items.Tables(0).Rows(i)("RoHS")
                                            CLMasterSQLCommand.Parameters("@InvoiceNo").Value = Header.InvoiceNo
                                            CLMasterSQLCommand.Parameters("@BillofLading").Value = Header.BillOfLading
                                            CLMasterSQLCommand.Parameters("@DN").Value = Header.DeliveryNote
                                            CLMasterSQLCommand.Parameters("@HeaderText").Value = Header.HeaderText
                                            CLMasterSQLCommand.Parameters("@SourceID").Value = ""
                                            If Header.Type = 3 AndAlso DRLine(0)("po_type") = "KB_PO" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 301
                                            ElseIf Header.Type = 3 AndAlso DRLine(0)("po_type") = "EJ_PO" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 302
                                                CLMasterSQLCommand.Parameters("@SourceID").Value = DRLine(0)("ejit_po_id")
                                            ElseIf Header.Type = 4 AndAlso DRLine(0)("ir_type") = "KB_IR" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 401
                                            ElseIf Header.Type = 4 AndAlso DRLine(0)("ir_type") = "EJ_IR" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 402
                                                CLMasterSQLCommand.Parameters("@SourceID").Value = DRLine(0)("ejit_ir_id")
                                                CLMasterSQLCommand.Parameters("@COO").Value = "CN"
                                            ElseIf Header.Type = 4 AndAlso DRLine(0)("ir_type") = "NO_IR" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 403
                                            ElseIf Header.Type = 7 AndAlso DRLine(0)("po_type") = "KB_ASNPO" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 701
                                            ElseIf Header.Type = 7 AndAlso DRLine(0)("po_type") = "EJIT_ASNPO" Then
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = 702
                                                CLMasterSQLCommand.Parameters("@SourceID").Value = DRLine(0)("ejit_po_id")

                                            Else
                                                CLMasterSQLCommand.Parameters("@ProcessID").Value = Header.Type
                                                If Header.Type = 9 Then                          'Identify which Receipt Type comes from here
                                                    Dim tmpID As Integer = CInt(Header.Type & "0" & Header.OrderType)
                                                    CLMasterSQLCommand.Parameters("@ProcessID").Value = tmpID
                                                End If

                                                'Record EJit ID for EJit RT              04/22/2014
                                                If myEJITID <> "" Then CLMasterSQLCommand.Parameters("@SourceID").Value = myEJITID
                                            End If
                                            CLMasterSQLCommand.Parameters("@ExpDate").Value = Items.Tables(0).Rows(i)("ExpDate")
                                            CLMasterSQLCommand.Parameters("@ProdDate").Value = Items.Tables(0).Rows(i)("ProductionDate")
                                            CLMasterSQLCommand.Parameters("@StockType").Value = Items.Tables(0).Rows(i)("StockType")
                                            CLMasterSQLCommand.Parameters("@MatSuffix1").Value = Items.Tables(0).Rows(i)("MatSuffix1")
                                            CLMasterSQLCommand.Parameters("@MatSuffix2").Value = Items.Tables(0).Rows(i)("MatSuffix2")
                                            CLMasterSQLCommand.Parameters("@MatSuffix3").Value = Items.Tables(0).Rows(i)("MatSuffix3")
                                            CLMasterSQLCommand.Parameters("@MaterialDesc").Value = Items.Tables(0).Rows(i)("MaterialDesc").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@VendorName").Value = Items.Tables(0).Rows(i)("VendorName").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@VendorPN").Value = Items.Tables(0).Rows(i)("VendorPart").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@StoragePosition").Value = Items.Tables(0).Rows(i)("StoragePosition").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@Operator").Value = Items.Tables(0).Rows(i)("Operator").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@IsTraceable").Value = Items.Tables(0).Rows(i)("Traceable").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@AddlText").Value = Items.Tables(0).Rows(i)("ItemText").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@Manufacturer").Value = Items.Tables(0).Rows(i)("Manufacturer").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@ManufacturerPN").Value = Items.Tables(0).Rows(i)("ManufacturerPN").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@QMLStatus").Value = Items.Tables(0).Rows(i)("QMLStatus").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@AddlData").Value = Items.Tables(0).Rows(i)("AddlData").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@Stemp").Value = Items.Tables(0).Rows(i)("Stemp").ToString.Trim
                                            CLMasterSQLCommand.Parameters("@MSL").Value = Items.Tables(0).Rows(i)("MSL").ToString.Trim

                                            'Save the Source CLID and record into T_CLMaster    "Yudy20140605
                                            If Header.OrderType = 4 Then CLMasterSQLCommand.Parameters("@SourceCLID").Value = SourceCLID

                                            'Change VendorID/VendorName to blank for OSP PO Receiving Charles 7/20/2018
                                            If Header.OrderType = 2 Then
                                                CLMasterSQLCommand.Parameters("@VendorID").Value = ""
                                                CLMasterSQLCommand.Parameters("@VendorName").Value = ""
                                            End If
                                            Try
                                                CLMasterSQLCommand.CommandType = CommandType.Text
                                                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                                                'ErrorLogging("Receiving-InsertSql", LoginData.User.ToUpper, "ReceiptNo: " & GRNo & CLMasterSQLCommand.CommandText, "I")
                                                ra = CLMasterSQLCommand.ExecuteNonQuery()
                                                DR(j)("POLineID") = TransactionID
                                                DR(j)("LID") = NextCLID
                                                DR(j).AcceptChanges()
                                            Catch ex As Exception
                                                ErrorLogging("Receiving-PostGR-InsertCLID", LoginData.User.ToUpper, "ReceiptNo: " & GRNo & " with Receipt type " & Header.Type & " and NextCLID " & NextCLID & "; " & ex.Message & ex.Source, "E")
                                            End Try
                                        End If
                                    Next

                                    Dim ShipInstr As String = ""
                                    If (Header.Type = 3 AndAlso DRLine(0)("po_type") = "EJ_PO") OrElse (Header.Type = 7 AndAlso DRLine(0)("po_type") = "EJIT_ASNPO") _
                                      OrElse (Header.Type = 4 AndAlso DRLine(0)("ir_type") = "EJ_IR") OrElse myEJITID <> "" Then
                                        'Update eJIT table
                                        Dim strJIT As String
                                        Dim cmdSQL As SqlClient.SqlCommand
                                        Dim eJITID, OrderNumber, Line, RecDate As String
                                        Dim JITQty As Decimal

                                        Try
                                            If myEJITID <> "" Then            'Record EJit ID for EJit RT              04/22/2014
                                                eJITID = myEJITID
                                            Else
                                                If Header.Type = 3 Then
                                                    eJITID = DRLine(0)("ejit_po_id")
                                                Else
                                                    eJITID = DRLine(0)("ejit_ir_id")
                                                End If
                                            End If
                                            OrderNumber = Items.Tables(0).Rows(i)("PONo").ToString
                                            Line = Items.Tables(0).Rows(i)("POItem").ToString
                                            JITQty = 0

                                            strJIT = String.Format("Select RecQty FROM T_KBPickingList with (nolock) WHERE EJITID='{0}'", eJITID)
                                            cmdSQL = New SqlClient.SqlCommand(strJIT, myConn)
                                            cmdSQL.CommandTimeout = TimeOut_M5

                                            If FixNull(cmdSQL.ExecuteScalar) <> "" Then
                                                JITQty = Convert.ToDecimal(cmdSQL.ExecuteScalar)
                                                JITQty = JITQty + Convert.ToDecimal(Items.Tables(0).Rows(i)("EntryQty"))
                                            Else
                                                JITQty = Convert.ToDecimal(Items.Tables(0).Rows(i)("EntryQty"))
                                            End If
                                            RecDate = Format(DateTime.Now, "MM/dd/yyyy").ToString
                                            strJIT = String.Format("UPDATE T_KBPickingList set RTNO='{1}', RecQty='{2}',RecBy = '{3}', RecOn='{4}', OrderNumber='{5}', LineNumber='{6}' where EJITID='{0}'", eJITID, GRNo, JITQty, LoginData.User.ToUpper, RecDate, OrderNumber, Line)
                                            cmdSQL = New SqlClient.SqlCommand(strJIT, myConn)
                                            cmdSQL.CommandTimeout = TimeOut_M5
                                            cmdSQL.ExecuteNonQuery()

                                            'Read Shipping Instruction for eJIT PO
                                            SqlStr = String.Format("Select ShipInstruction FROM T_KBPickingList with (nolock) WHERE EJITID='{0}'", eJITID)
                                            Sqlcmd = New SqlClient.SqlCommand(SqlStr, myConn)
                                            ShipInstr = FixNull(Sqlcmd.ExecuteScalar)
                                        Catch ex As Exception
                                            ErrorLogging("Receiving-PostGR-Update_eJIT", LoginData.User.ToUpper, "ReceiptNo: " & GRNo & " with Receipt type " & Header.Type & " and EJITID " & eJITID & "; " & ex.Message & ex.Source, "E")
                                        End Try
                                    End If


                                    If Header.OrderType <> 2 Then

                                        'Record the Pallet info which needs to do Component Delivery
                                        If RecConfig = "YES" AndAlso PORoutingID < 3 Then
                                            Dim DBdr() As DataRow = Nothing
                                            DBdr = dtDashB.Select(" MaterialNo = '" & Items.Tables(0).Rows(i)("MaterialNo").ToString & "'")
                                            If DBdr.Length = 0 Then
                                                Dim myDR As DataRow
                                                myDR = dtDashB.NewRow()
                                                myDR("MaterialNo") = Items.Tables(0).Rows(i)("MaterialNo").ToString
                                                myDR("Packages") = NoOfPackage
                                                myDR("RecQty") = Items.Tables(0).Rows(i)("EntryQty")
                                                myDR("RecRouting") = Items.Tables(0).Rows(i)("StockType").ToString
                                                myDR("PalletID") = Items.Tables(0).Rows(i)("StoragePosition").ToString
                                                dtDashB.Rows.Add(myDR)
                                            Else
                                                DBdr(0)("Packages") = DBdr(0)("Packages") + NoOfPackage
                                                DBdr(0)("RecQty") = DBdr(0)("RecQty") + Items.Tables(0).Rows(i)("EntryQty")
                                            End If
                                        End If


                                        'Write RT Slip data into eTrace database for printing use
                                        Dim LabelData As RTLabel
                                        LabelData.RT = GRNo
                                        LabelData.OrgCode = LoginData.OrgCode
                                        LabelData.RecDate = PostDate.ToString        'Header.PostDate.ToString("MM/dd/yyyy")

                                        LabelData.OrgName = DRLine(0)("Organization_name").ToString
                                        LabelData.Buyer = DRLine(0)("employee_number").ToString
                                        LabelData.BuyerName = DRLine(0)("buyer").ToString

                                        'For the RT which need to do AML verification, will print signal on the RT Slip   -- 02/16/2017
                                        LabelData.Vendor = Items.Tables(0).Rows(i)("VendorNo")
                                        If myStatusCode = "9" Then LabelData.Vendor = Items.Tables(0).Rows(i)("VendorNo") & " ( *V* ) "

                                        LabelData.VendorName = Items.Tables(0).Rows(i)("VendorName")
                                        LabelData.InvoiceDN = Header.InvoiceNo.ToString
                                        LabelData.PONo = Items.Tables(0).Rows(i)("PONo")
                                        LabelData.POLine = 1
                                        LabelData.Shipment = 1
                                        LabelData.POType = ""
                                        If Header.OrderType <> 4 Then LabelData.POType = DRLine(0)("po_header_type").ToString

                                        'Combine Shipping Instruction to PO Type for eJIT PO  -- 3/24/2015
                                        If Header.OrderType = 3 AndAlso ShipInstr <> "" Then
                                            LabelData.POType = LabelData.POType & " / " & ShipInstr
                                        End If

                                        'Add Vendor Consignment Flag to RT Label      12/06/2013
                                        LabelData.VMIFlag = ""
                                        If Header.OrderType = 1 AndAlso DRLine(0)("vmi_flag").ToString = "Y" Then LabelData.VMIFlag = "CVMI"

                                        Dim POLine() As String
                                        POLine = Split(Items.Tables(0).Rows(i)("POItem"), ".")

                                        If POLine Is DBNull.Value OrElse POLine.Length = 0 Then
                                        ElseIf POLine.Length = 1 Then
                                            LabelData.POLine = POLine(0)
                                        ElseIf POLine.Length >= 2 Then
                                            LabelData.POLine = POLine(0)
                                            LabelData.Shipment = POLine(1)
                                        End If

                                        'LabelData.PORev = Items.Tables(0).Rows(i)("MaterialRevision").ToString
                                        LabelData.Material = Items.Tables(0).Rows(i)("MaterialNo").ToString
                                        LabelData.MatDesc = Items.Tables(0).Rows(i)("MaterialDesc").ToString
                                        LabelData.RoHS = Items.Tables(0).Rows(i)("RoHS").ToString

                                        Dim RTQty As Decimal = CDec(Items.Tables(0).Rows(i)("EntryQty"))
                                        LabelData.Qty = Format(RTQty, "#0.#####")               'Remove the decimals
                                        'LabelData.Qty = Items.Tables(0).Rows(i)("EntryQty").ToString

                                        LabelData.UOM = Items.Tables(0).Rows(i)("BaseUOM").ToString
                                        LabelData.NoOfPackage = NoOfPackage.ToString
                                        LabelData.AddlText = Items.Tables(0).Rows(i)("ItemText").ToString
                                        LabelData.TempLoc = Items.Tables(0).Rows(i)("StoragePosition").ToString
                                        LabelData.TransactBy = LoginData.User.ToUpper
                                        LabelData.OptrName = Items.Tables(0).Rows(i)("Operator").ToString

                                        LabelData.Safety = ""
                                        Dim AddlData As String = Items.Tables(0).Rows(i)("AddlData").ToString
                                        If InStr(AddlData, "S") > 0 Then LabelData.Safety = "Safety"

                                        LabelData.Mfr = Items.Tables(0).Rows(i)("Manufacturer").ToString
                                        LabelData.MfrPN = Items.Tables(0).Rows(i)("ManufacturerPN").ToString
                                        LabelData.QMLStatus = Items.Tables(0).Rows(i)("QMLStatus").ToString

                                        LabelData.ExpDate = ""
                                        If Not Items.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
                                            LabelData.ExpDate = CDate(Items.Tables(0).Rows(i)("ExpDate")).ToShortDateString
                                        End If

                                        LabelData.InspFlag = Items.Tables(0).Rows(i)("StockType").ToString
                                        LabelData.ConFactor = Items.Tables(0).Rows(i)("BaseUOMConv").ToString

                                        'For Normal PO or Normal IR get PredefinedSubInv/PredefinedLocator, others get SubInv/Locator  "Yudy20110428
                                        'Please note: for Header.Type = 9 and Header.OrderType = 4, no column "ir_type" in table po_line    "12/13/2013
                                        If Header.OrderType = 1 OrElse (Header.Type = 4 AndAlso DRLine(0)("ir_type") = "NO_IR") OrElse Header.Type = 6 OrElse Header.Type = 7 Then
                                            LabelData.DestSubInv = Items.Tables(0).Rows(i)("PredefinedSubInv").ToString
                                            LabelData.DestLocator = Items.Tables(0).Rows(i)("PredefinedLocator").ToString
                                        Else
                                            LabelData.DestSubInv = Items.Tables(0).Rows(i)("SubInventory").ToString
                                            LabelData.DestLocator = Items.Tables(0).Rows(i)("Locator").ToString
                                        End If
                                        LabelData.Stemp = Items.Tables(0).Rows(i)("Stemp").ToString
                                        LabelData.MSL = Items.Tables(0).Rows(i)("MSL").ToString
                                        LabelData.COC = Items.Tables(0).Rows(i)("COC").ToString

                                        LabelData.MCPosition = MCPosition
                                        LabelData.StorageType = DRLine(0)("storage_type").ToString
                                        'LabelData.SysRev = DRLine(0)("item_rev").ToString
                                        'LabelData.FAIStatus = DRLine(0)("fai_rep_num").ToString                      'Get FAI Status
                                        LabelData.FAIStatus = Items.Tables(0).Rows(i)("EER").ToString               'Add EER from iPro to RT Slip   --10/12/2016

                                        'Get PO Item Revision and iPro Revision here                                                   '05/04/2017
                                        LabelData.PORev = DRLine(0)("po_item_reversion").ToString
                                        LabelData.SysRev = ""
                                        LabelData.ESD = ""
                                        LabelData.ShelfLife = ""

                                        Dim RESData As String = Items.Tables(0).Rows(i)("RESData").ToString
                                        If RESData <> "" Then                                                                                           '05/04/2017
                                            Dim Arry() As String = Split(RESData, "~")
                                            If Arry.Length > 0 Then
                                                LabelData.SysRev = Arry(0)                          'Get iPro Revision
                                                LabelData.ESD = Arry(1)                               'Get iPr ESD Flag
                                                LabelData.ShelfLife = Arry(2)                     'Get iPro ShelfLife Days
                                            End If
                                        End If


                                        'Get Replenishment Priority and Sample Size from RT return message
                                        LabelData.Priority = ""
                                        LabelData.Size = ""
                                        If Header.Type <> 9 Then
                                            LabelData.Priority = msgDR(0)("Priority").ToString
                                            If (Header.Type = 1 Or Header.Type = 6) AndAlso GRResult.Tables(0).Rows.Count > 0 Then
                                                Dim myDR() As DataRow = Nothing
                                                myDR = GRResult.Tables(0).Select(" msg_type = 'SAMPLE_SIZE' and msg_id = '" & Items.Tables(0).Rows(i)("POLineID") & "'")
                                                If myDR.Length > 0 Then LabelData.Size = myDR(0)("text").ToString
                                            End If
                                        End If

                                        LabelData.MatDesc = Replace(LabelData.MatDesc, "^", "~")
                                        LabelData.MfrPN = Replace(LabelData.MfrPN, "^", "~")

                                        'Combine DateCode & LotNo, and print it out on RT Slip which requested by PH  -- 05/16/2014
                                        LabelData.DateCode = Replace(DateCode, "^", "~")
                                        LabelData.LotNo = Replace(LotNo, "^", "~")

                                        If PORoutingID = 2 Then
                                            LabelData.HisInsRT = InspectionRT
                                        Else
                                            LabelData.HisInsRT = ""
                                        End If
                                        Try
                                            Dim da As DataAccess = GetDataAccess()
                                            Dim SqlstrSampleloc As String
                                            Dim VSampleLoc As String
                                            Dim SLDataRow As Data.DataRow
                                            drSampleloc = Nothing
                                            drSampleloc = dsSampleLoc.Tables(0).Select(" OrgCode = '" & LoginData.OrgCode & "' and Item = '" & Items.Tables(0).Rows(i)("MaterialNo") & "'")
                                            If drSampleloc.Length > 0 Then
                                                LabelData.SampleLoc = drSampleloc(0)("SampleLocation").ToString
                                            Else
                                                SqlstrSampleloc = String.Format("exec dbo.ora_Sample_Location '{0}','{1}'", LoginData.OrgID.ToString, Items.Tables(0).Rows(i)("MaterialNo"))
                                                VSampleLoc = Convert.ToString(da.ExecuteScalar(SqlstrSampleloc))
                                                LabelData.SampleLoc = VSampleLoc
                                                SLDataRow = dsSampleLoc.Tables(0).NewRow
                                                SLDataRow("OrgCode") = LoginData.OrgCode
                                                SLDataRow("Item") = Items.Tables(0).Rows(i)("MaterialNo")
                                                SLDataRow("SampleLocation") = VSampleLoc
                                                dsSampleLoc.Tables(0).Rows.Add(SLDataRow)
                                                dsSampleLoc.AcceptChanges()
                                            End If

                                            Dim TranID As String = ""
                                            If Header.Type = 9 Then
                                                'Check if the same RTSlip already record in T_RTSlip or not, if yes, no need to save again
                                                SqlStr = String.Format("Select TransactionID from T_RTSlip with (nolock) where TransactionID = '{0}' ", TransactionID)
                                                TranID = Convert.ToString(da.ExecuteScalar(SqlStr))
                                            End If

                                            If TranID = "" Then
                                                Dim RTData As String   'LabelData.POType
                                                RTData = "RT^" & LabelData.RT & "^RecDate^" & LabelData.RecDate & "^Priority^" & LabelData.Priority & "^DestSubInv^" & LabelData.DestSubInv & "^DestLocator^" & LabelData.DestLocator & "^MCPosition^" & LabelData.MCPosition & "^StorageType^" & LabelData.StorageType _
                                                        & "^OrgCode^" & LabelData.OrgCode & "^OrgName^" & LabelData.OrgName & "^Vendor^" & LabelData.Vendor & "^VendorName^" & LabelData.VendorName & "^InvoiceDN^" & LabelData.InvoiceDN & "^PONo^" & LabelData.PONo & "^POLine^" & LabelData.POLine _
                                                        & "^Shipment^" & LabelData.Shipment & "^VMIFlag^" & LabelData.VMIFlag & "^POType^" & LabelData.POType & "^Buyer^" & LabelData.Buyer & "^BuyerName^" & LabelData.BuyerName & "^SysRev^" & LabelData.SysRev & "^PORev^" & LabelData.PORev & "^Material^" & LabelData.Material _
                                                        & "^Desc^" & LabelData.MatDesc & "^Qty^" & LabelData.Qty & "^UOM^" & LabelData.UOM & "^NoOfPackage^" & LabelData.NoOfPackage & "^AddlText^" & LabelData.AddlText & "^RoHS^" & LabelData.RoHS & "^TempLoc^" & LabelData.TempLoc _
                                                        & "^ConFactor^" & LabelData.ConFactor & "^ExpDate^" & LabelData.ExpDate & "^InspFlag^" & LabelData.InspFlag & "^FAIStatus^" & LabelData.FAIStatus & "^Safety^" & LabelData.Safety & "^Size^" & LabelData.Size _
                                                        & "^MSL^" & LabelData.MSL & "^Stemp^" & LabelData.Stemp & "^COC^" & LabelData.COC & "^TransactBy^" & LabelData.TransactBy & "^OptrName^" & LabelData.OptrName _
                                                        & "^Mfr^" & LabelData.Mfr & "^MfrPN^" & LabelData.MfrPN & "^QMLStatus^" & LabelData.QMLStatus & "^DateCode^" & LabelData.DateCode & "^LotNo^" & LabelData.LotNo _
                                                        & "^ESD^" & LabelData.ESD & "^ShelfLife^" & LabelData.ShelfLife & "^HisInsRT^" & LabelData.HisInsRT & "^SampleLoc^" & LabelData.SampleLoc

                                                'Filter Special Characters for RTData
                                                RTData = FilterSpecial(RTData)

                                                'Filter Special Characters for MFR / MPN
                                                Dim MFR, MPN As String
                                                MFR = Items.Tables(0).Rows(i)("Manufacturer").ToString
                                                MPN = Items.Tables(0).Rows(i)("ManufacturerPN").ToString
                                                MFR = FilterSpecial(MFR)
                                                MPN = FilterSpecial(MPN)

                                                SqlStr = String.Format("INSERT INTO T_RTSlip (TransactionID,OrgCode,RTNo,MaterialNo,PONo,POLine,RecDate,RecQty,MFR,MPN,AMLStatus,RTContent,CreatedOn,CreatedBy) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}', getDate(),'{12}')", _
                                                         TransactionID, LoginData.OrgCode, GRNo, Items.Tables(0).Rows(i)("MaterialNo").ToString, Items.Tables(0).Rows(i)("PONo"), Items.Tables(0).Rows(i)("POItem"), PostDate, Items.Tables(0).Rows(i)("EntryQty"), MFR, MPN, Items.Tables(0).Rows(i)("QMLStatus").ToString, RTData, LoginData.User.ToUpper)
                                                da.ExecuteNonQuery(SqlStr)
                                            End If

                                        Catch ex As Exception
                                            ErrorLogging("Receiving-PostGR-RecordRTData", LoginData.User.ToUpper, "ReceiptNo: " & GRNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                                        End Try

                                    End If
                                End If

                                'Save Receiving TransactionID as POLineID for RT Slip reprint use in client
                                Items.Tables(0).Rows(i)("POLineID") = TransactionID
                            End If
                        End If

                    End If

                Next
                Items.AcceptChanges()
                If CLIDVerifyFlag = True Then
                    PostGR.GRStatus = "Y/V"        'Set Flag to record the RT which need to do AML Verification
                End If

                'Save the Pallet info in table T_RecDashboard which needs to do Component Delivery
                If Header.OrderType <> 2 Then
                    If RecConfig = "YES" AndAlso dtDashB.Rows.Count > 0 Then
                        Dim Status As String = "10"           'New Receipt
                        For i = 0 To dtDashB.Rows.Count - 1
                            SqlStr = String.Format("INSERT INTO T_RecDashboard (OrgCode,RTNo,MaterialNo,CreatedOn,CreatedBy,Packages,RecQty,RecRouting,PalletID,Status,StaChangedOn,StaChangedBy) values ('{0}','{1}','{2}',getDate(),'{3}','{4}','{5}','{6}','{7}','{8}',getDate(),'{9}')", _
                                     LoginData.OrgCode, GRNo, dtDashB.Rows(i)("MaterialNo"), LoginData.User.ToUpper, dtDashB.Rows(i)("Packages"), dtDashB.Rows(i)("RecQty"), dtDashB.Rows(i)("RecRouting"), dtDashB.Rows(i)("PalletID"), Status, LoginData.User.ToUpper)
                            Sqlcmd = New SqlClient.SqlCommand(SqlStr, myConn)
                            Sqlcmd.CommandTimeout = TimeOut_M5
                            Sqlcmd.ExecuteNonQuery()
                        Next
                    End If
                End If

                PostGR.CLIDs = New DataSet
                PostGR.CLIDs = Items
                myConn.Close()

            Catch ex As Exception
                PostGR.GRNo = ""
                PostGR.GRStatus = "E"
                PostGR.GRMessage = "Receipt Error"
                ErrorLogging("Receiving-PostGR1", LoginData.User.ToUpper, "ReceiptNo: " & GRNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
            Finally
                If myConn.State <> ConnectionState.Closed Then myConn.Close()
            End Try
            If PostGR.GRStatus = "E" Then Exit Function


            Dim IRMergeFlag As Boolean = False
            If Header.OrderType = 4 AndAlso IRLabels.Tables("CLIDs").Rows.Count > 0 Then       'Update RTNo to Source Server for ISO
                Dim myIRData As IRData = New IRData
                myIRData.User = LoginData.User.ToUpper
                myIRData.SourceOrg = PODetails.Tables("po_header").Rows(0)("vendor_number").ToString
                myIRData.DestOrg = LoginData.OrgCode
                myIRData.ShipmentNo = Header.OrderNo
                myIRData.IRRTNo = GRNo

                Dim UpdateIRNo As Boolean
                UpdateIRNo = UpdateRTNoToServer(myIRData, IRLabels)

                'Check if there has merged CLID from shipments
                DR = Items.Tables(1).Select("RefCLID <> ''", "RefCLID DESC")
                If DR.Length > 0 Then
                    If MergePackLabels(DR, "M", "RefCLID", PrintLabels) = True Then
                        Items.AcceptChanges()
                    End If
                    IRMergeFlag = True
                End If
            End If

            If Header.OrderType <> 2 Then
                'Now merge those marked labels
                If IRMergeFlag = False Then
                    'DR = Items.Tables(1).Select("RefCLID >= 'M1'", "RefCLID DESC")
                    DR = Items.Tables(1).Select("RefCLID >= 'M001'", "RefCLID DESC")
                    If DR.Length > 0 Then
                        If MergePackLabels(DR, "M", "RefCLID") = True Then
                            Items.AcceptChanges()
                        End If
                    End If
                End If

                'Now generate BoxIDs
                DR = Items.Tables(1).Select("BoxID >= 'B001'", "BoxID DESC")
                If DR.Length > 0 Then
                    If MergePackLabels(DR, "P", "BoxID") = True Then
                        Items.AcceptChanges()
                    End If
                End If
            End If

        Catch ex As Exception
            ErrorLogging("Receiving-PostGR", LoginData.User.ToUpper, "OrderNo " & Header.OrderNo & ", " & ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Private Function MergePackLabels(ByRef DR() As DataRow, ByVal Action As String, ByVal FieldName As String, Optional ByVal CreateNewCLID As Boolean = True) As Boolean
        Dim Labels As DataSet = New DataSet
        Dim CLIDsTable As DataTable
        Dim myDataColumn As DataColumn
        CLIDsTable = New Data.DataTable("CLIDTable")
        myDataColumn = New Data.DataColumn("CLID", System.Type.GetType("System.String"))
        CLIDsTable.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("RefCLID", System.Type.GetType("System.String"))
        CLIDsTable.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("StockType", System.Type.GetType("System.String"))
        CLIDsTable.Columns.Add(myDataColumn)

        Labels.Tables.Add(CLIDsTable)

        Dim myDataRow As Data.DataRow
        Dim StartingRow As Integer = 0
        Dim ReturnedID As String
        Dim i, j As Integer

        Try

            For i = 0 To DR.Length - 1
                If DR(i)(FieldName) <> DR(StartingRow)(FieldName) Then
                    If Action = "P" Then   ' Pack
                        ReturnedID = PackLabels(Labels, CreateNewCLID)
                    Else                   ' Merge
                        ReturnedID = LabelMerge(Labels, CreateNewCLID)
                    End If
                    If Microsoft.VisualBasic.Left(ReturnedID, 5) = "ERROR" Then
                        MergePackLabels = False
                        Exit Function
                    End If
                    For j = StartingRow To i - 1
                        DR(j)(FieldName) = ReturnedID
                    Next
                    StartingRow = i
                    Labels.Clear()
                End If

                'Find the Inspection Status for Valid CLIDs for later use
                Dim SqlStr, StockType As String
                Dim da As DataAccess = GetDataAccess()
                SqlStr = String.Format("Select StockType from T_CLMaster with (nolock) where CLID = '{0}' ", DR(i)("LID").ToString)
                StockType = Convert.ToString(da.ExecuteScalar(SqlStr))

                ' For merged IDs, need to find the valid IDs
                myDataRow = Labels.Tables("CLIDTable").NewRow()
                myDataRow(0) = DR(i)("LID").ToString.Trim
                myDataRow(1) = DR(i)("RefCLID").ToString.Trim
                myDataRow(2) = StockType

                If Action = "P" Then
                    If Not DR(i)("RefCLID") Is DBNull.Value Then
                        If Not DR(i)("RefCLID") = "" Then
                            myDataRow(0) = DR(i)("RefCLID")
                        End If
                    End If
                End If
                Labels.Tables("CLIDTable").Rows.Add(myDataRow)
            Next
            If Action = "P" Then    ' Pack
                ReturnedID = PackLabels(Labels, CreateNewCLID)
            Else                    ' Merge
                ReturnedID = LabelMerge(Labels, CreateNewCLID)
            End If
            If Microsoft.VisualBasic.Left(ReturnedID, 5) = "ERROR" Then
                MergePackLabels = False
            Else
                For j = StartingRow To DR.Length - 1
                    DR(j)(FieldName) = ReturnedID
                Next
                MergePackLabels = True
            End If

        Catch ex As Exception
            ErrorLogging("Receiving-MergePackLabels", "", ex.Message & ex.Source, "E")
            MergePackLabels = False
        End Try

    End Function

    Private Function LabelMerge(ByVal Labels As DataSet, Optional ByVal CreateNewCLID As Boolean = True) As String

        If Labels.Tables(0).Rows.Count < 2 Then Return ""

        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1 As SqlClient.SqlCommand
        Dim CLMasterSQLCommand2 As SqlClient.SqlCommand
        Dim CLMasterSQLCommand3 As SqlClient.SqlCommand
        Dim myCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim NextCLID As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim i, k As Integer
        Dim sum1 As Decimal
        Dim sum2 As Decimal

        If CreateNewCLID = True Then
            myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.AddWithValue("@NextNo", "")
            myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
            myCommand.Parameters(0).Size = 20
            myCommand.Parameters(0).Direction = ParameterDirection.Output
            myCommand.Parameters.AddWithValue("@TypeID", "CLID")
            myCommand.CommandTimeout = TimeOut_M5
        End If

        Try
            Dim DR() As DataRow
            Dim strCondition As String = Labels.Tables(0).Rows(0)(0).ToString
            DR = Labels.Tables(0).Select("StockType <> 'FTS' ")
            If DR.Length > 0 Then strCondition = DR(0)(0).ToString
            If strCondition = "" Then Return ""

            myConn.Open()

            NextCLID = ""
            If CreateNewCLID = False Then
                NextCLID = Labels.Tables(0).Rows(0)(1).ToString         'RefCLID
            Else
                'Try up to 5 times when failed getting next id
                k = 0
                While (k < 5 And NextCLID = "")
                    Try
                        ra = myCommand.ExecuteNonQuery()
                        NextCLID = myCommand.Parameters(0).Value
                    Catch ex As Exception
                        k = k + 1
                        ErrorLogging("Receiving-LabelMerge", "Deadlocked? " & Str(k), "Failed getting next ID; " & ex.Message & ex.Source, "E")
                    End Try
                End While
            End If
            If NextCLID = "" Then
                LabelMerge = "Error"
                ErrorLogging("Receiving-LabelMerge", "", Str(k) & " Failed getting next ID; Original LabelID: " & Labels.Tables(0).Rows(0)(0), "I")
                Exit Function
            End If

            Try
                CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_CLMaster (CLID, OrgCode, Qty, QtyBaseUOM, MaterialNo, MaterialRevision, MaterialDesc, UOM, BaseUOM, CreatedOn, CreatedBy, ChangedOn, DateCode, LotNo, CountryOfOrigin, StatusCode,ExpDate,RecDocNo,RTLot,RecDate,ProcessID,RoHS,PurOrdNo,PurOrdItem,DeliveryType,VendorID,VendorName,VendorPN,InvoiceNo,BillofLading,DN,HeaderText,ProdDate,ReasonCode,StoragePosition,PredefinedSubInv,PredefinedLocator,Operator,StockType,ItemText,MatSuffix1,MatSuffix2,MatSuffix3,Printed,IsTraceable,AddlText,MatDocNo,BoxID,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,SLOC,StorageBin,SourceID,LastTransaction,SourceCLID,MCPosition ) SELECT '" & NextCLID & "', OrgCode, 0, 0, MaterialNo, MaterialRevision, MaterialDesc, UOM, BaseUOM, CreatedOn, CreatedBy, ChangedOn, DateCode, LotNo, CountryOfOrigin, StatusCode,ExpDate,RecDocNo,RTLot,RecDate,ProcessID,RoHS,PurOrdNo,PurOrdItem,DeliveryType,VendorID,VendorName,VendorPN,InvoiceNo,BillofLading,DN,HeaderText,ProdDate,ReasonCode,StoragePosition,PredefinedSubInv,PredefinedLocator,Operator,StockType,ItemText,MatSuffix1,MatSuffix2,MatSuffix3,Printed,IsTraceable,AddlText,MatDocNo,BoxID,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,SLOC,StorageBin,SourceID,LastTransaction,SourceCLID,MCPosition FROM T_CLMaster AS T_CLMaster_1 with (nolock) WHERE (CLID = '" & strCondition & "')", myConn)
                CLMasterSQLCommand.CommandType = CommandType.Text
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand.ExecuteNonQuery()
            Catch ex As Exception
                LabelMerge = "Error"
                ErrorLogging("Receiving-LabelMerge-InsertCLID", "", "NextCLID " & NextCLID & "; " & ex.Message & ex.Source, "E")
                Exit Function
            End Try

            CLMasterSQLCommand1 = New SqlClient.SqlCommand("select Qty,QtyBaseUOM from T_CLMaster with (nolock) where CLID=@CLID", myConn)
            CLMasterSQLCommand1.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")

            CLMasterSQLCommand3 = New SqlClient.SqlCommand("update T_CLMaster set StatusCode=0,ReferenceCLID=@ReferenceCLID where CLID=@CLID", myConn)
            CLMasterSQLCommand3.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
            CLMasterSQLCommand3.Parameters.Add("@ReferenceCLID", SqlDbType.VarChar, 20, "ReferenceCLID")


            For i = 0 To Labels.Tables(0).Rows.Count - 1
                CLMasterSQLCommand1.Parameters("@CLID").Value = Labels.Tables(0).Rows(i)("CLID")
                CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
                objReader = CLMasterSQLCommand1.ExecuteReader()
                While objReader.Read()
                    sum1 = sum1 + objReader.GetValue(0)
                    sum2 = sum2 + objReader.GetValue(1)
                End While
                objReader.Close()

                CLMasterSQLCommand3.Parameters("@CLID").Value = Labels.Tables(0).Rows(i)("CLID")
                CLMasterSQLCommand3.Parameters("@ReferenceCLID").Value = NextCLID
                CLMasterSQLCommand3.CommandType = CommandType.Text
                CLMasterSQLCommand3.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand3.ExecuteNonQuery()
            Next
            CLMasterSQLCommand2 = New SqlClient.SqlCommand("update T_CLMaster set Qty=@Qty,QtyBaseUOM=@QtyBaseUOM WHERE (CLID = '" & NextCLID & "')", myConn)
            CLMasterSQLCommand2.Parameters.Add("@Qty", SqlDbType.Float, 20, "Qty")
            CLMasterSQLCommand2.Parameters.Add("@QtyBaseUOM", SqlDbType.Float, 20, "QtyBaseUOM")
            CLMasterSQLCommand2.Parameters("@Qty").Value = sum1
            CLMasterSQLCommand2.Parameters("@QtyBaseUOM").Value = sum2

            CLMasterSQLCommand2.CommandType = CommandType.Text
            CLMasterSQLCommand2.CommandTimeout = TimeOut_M5
            ra = CLMasterSQLCommand2.ExecuteNonQuery()

            LabelMerge = NextCLID
            myConn.Close()
        Catch ex As Exception
            LabelMerge = "Error"
            ErrorLogging("Receiving-LabelMerge", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try


    End Function

    Private Function PackLabels(ByVal CLIDs As DataSet, Optional ByVal CreateNewCLID As Boolean = True) As String
        Dim myCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim NextBoxID As String
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim i, k As Integer

        myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.Parameters.AddWithValue("@NextNo", "")
        myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
        myCommand.Parameters(0).Size = 20
        myCommand.Parameters(0).Direction = ParameterDirection.Output
        myCommand.Parameters.AddWithValue("@TypeID", "BoxID")
        myCommand.CommandTimeout = TimeOut_M5

        Try
            myConn.Open()

            'Try up to 5 times when failed getting next id
            NextBoxID = ""
            k = 0
            While (k < 5 And NextBoxID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextBoxID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("Receiving-PackLabels", "Deadlocked? " & Str(k), "Failed getting next ID; " & ex.Message & ex.Source, "E")
                End Try
            End While
            If NextBoxID = "" Then
                PackLabels = "Error"
                ErrorLogging("Receiving-PackLabels", "", Str(k) & " Failed getting next ID; Original LabelID: " & CLIDs.Tables(0).Rows(0)("CLID"), "I")
                Exit Function
            End If

            CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set BoxID=@BoxID where CLID=@CLID ", myConn)
            CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
            CLMasterSQLCommand.Parameters.Add("@BoxID", SqlDbType.VarChar, 20, "BoxID")
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                CLMasterSQLCommand.Parameters("@CLID").Value = CLIDs.Tables(0).Rows(i)("CLID")
                CLMasterSQLCommand.Parameters("@BoxID").Value = NextBoxID
                CLMasterSQLCommand.CommandType = CommandType.Text
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand.ExecuteNonQuery()
            Next
            myConn.Close()
            PackLabels = NextBoxID
        Catch ex As Exception
            PackLabels = "Error"
            ErrorLogging("Receiving-PackLabels", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function SaveDCodeLotNo(ByVal LoginData As ERPLogin, ByVal CLIDs As DataSet) As DataSet
        Dim i As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer

        Dim UpdateAML As String = CLIDs.Tables(0).Rows(0)("UpdateAML").ToString.ToUpper

        If UpdateAML = "FALSE" Then
            CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set ChangedOn=getDate(), LastTransaction='Receiving-SaveDC/L', ChangedBy=@ChangedBy, DateCode=@DateCode,LotNo=@LotNo,CountryOfOrigin=@COO, MatSuffix1=@MatSuffix1,MatSuffix2=@MatSuffix2,MatSuffix3=@MatSuffix3 where CLID=@CLID ", myConn)
            CLMasterSQLCommand.Parameters.Add("@DateCode", SqlDbType.VarChar, 20, "DateCode")
            CLMasterSQLCommand.Parameters.Add("@LotNo", SqlDbType.VarChar, 50, "LotNo")
            CLMasterSQLCommand.Parameters.Add("@COO", SqlDbType.VarChar, 20, "CountryOfOrigin")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix1", SqlDbType.VarChar, 10, "MatSuffix1")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix2", SqlDbType.VarChar, 10, "MatSuffix2")
            CLMasterSQLCommand.Parameters.Add("@MatSuffix3", SqlDbType.VarChar, 10, "MatSuffix3")
        Else
            CLMasterSQLCommand = New SqlClient.SqlCommand("update T_CLMaster set ChangedOn=getDate(), LastTransaction='Receiving-UpdateAML', ChangedBy=@ChangedBy, Manufacturer=@Manufacturer,ManufacturerPN=@ManufacturerPN,QMLStatus=@QMLStatus where CLID=@CLID ", myConn)
            CLMasterSQLCommand.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 50, "Manufacturer")
            CLMasterSQLCommand.Parameters.Add("@ManufacturerPN", SqlDbType.VarChar, 50, "ManufacturerPN")
            CLMasterSQLCommand.Parameters.Add("@QMLStatus", SqlDbType.VarChar, 50, "QMLStatus")
        End If
        CLMasterSQLCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        CLMasterSQLCommand.Parameters.Add("@ChangedBy", SqlDbType.VarChar, 20, "ChangedBy")
        CLMasterSQLCommand.Parameters("@ChangedBy").Value = LoginData.User.ToUpper

        Try
            myConn.Open()
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                CLMasterSQLCommand.Parameters("@CLID").Value = CLIDs.Tables(0).Rows(i)("LID")

                If UpdateAML = "FALSE" Then
                    CLMasterSQLCommand.Parameters("@DateCode").Value = CLIDs.Tables(0).Rows(i)("DateCode")
                    CLMasterSQLCommand.Parameters("@LotNo").Value = CLIDs.Tables(0).Rows(i)("LotNo")
                    CLMasterSQLCommand.Parameters("@COO").Value = CLIDs.Tables(0).Rows(i)("COO")
                    CLMasterSQLCommand.Parameters("@MatSuffix1").Value = CLIDs.Tables(0).Rows(i)("MatSuffix1")
                    CLMasterSQLCommand.Parameters("@MatSuffix2").Value = CLIDs.Tables(0).Rows(i)("MatSuffix2")
                    CLMasterSQLCommand.Parameters("@MatSuffix3").Value = CLIDs.Tables(0).Rows(i)("MatSuffix3")
                Else
                    CLMasterSQLCommand.Parameters("@Manufacturer").Value = CLIDs.Tables(0).Rows(i)("Manufacturer")
                    CLMasterSQLCommand.Parameters("@ManufacturerPN").Value = CLIDs.Tables(0).Rows(i)("ManufacturerPN")
                    CLMasterSQLCommand.Parameters("@QMLStatus").Value = CLIDs.Tables(0).Rows(i)("QMLStatus")
                End If

                CLMasterSQLCommand.CommandType = CommandType.Text
                CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand.ExecuteNonQuery()
            Next
            myConn.Close()

            If UpdateAML = "FALSE" Then
                Dim DR() As DataRow
                DR = CLIDs.Tables(0).Select("BoxID LIKE 'B%'", "BoxID DESC")
                If DR.Length > 0 Then
                    If MergePackLabels(DR, "P", "BoxID") = True Then
                        CLIDs.AcceptChanges()
                    End If
                End If

                DR = CLIDs.Tables(0).Select("RefCLID LIKE 'M%'", "RefCLID DESC")
                If DR.Length > 0 Then
                    If MergePackLabels(DR, "M", "RefCLID") = True Then
                        CLIDs.AcceptChanges()
                    End If
                End If
            End If

            SaveDCodeLotNo = CLIDs

        Catch ex As Exception
            ErrorLogging("Receiving-SaveDCodeLotNo", LoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Return CLIDs
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function PrintRECLabels(ByVal CLIDs As DataSet, ByVal LabelPrinter As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            PrintRECLabels = True

            Try
                Dim dtItem As New DataTable
                dtItem.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
                dtItem.Columns.Add(New Data.DataColumn("PNSeq", System.Type.GetType("System.String")))

                Dim DR() As DataRow
                Dim myDR As DataRow
                Dim SeqFlag As Boolean = False

                Dim dtCLID As DataTable = New DataTable
                dtCLID = CLIDs.Tables(0).Clone

                Dim i As Integer
                Dim j As Integer = 0
                For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                    Dim CLID As String = CLIDs.Tables(0).Rows(i)(0).ToString
                    If CLID = "" Then Continue For

                    Dim StartStr, MidStr, LblStart As String
                    StartStr = Microsoft.VisualBasic.Left(CLID, 1)
                    MidStr = Microsoft.VisualBasic.Mid(CLID, 3, 1)
                    LblStart = Microsoft.VisualBasic.Left(CLID, 2)                     'LD Huawei requirement   -- 11/16/2016

                    'For Product Label printing, just call normal function to print it out
                    If StartStr = "B" OrElse MidStr = "P" OrElse LblStart = "LE" Then
                        If PrintCLID(CLID, LabelPrinter) = False Then
                            PrintRECLabels = False
                        End If
                        Sleep(5)
                        Continue For
                    Else

                        'Collect Material CLID for SeqNo printing later, this is only for Material Label Format printing
                        dtCLID.ImportRow(CLIDs.Tables(0).Rows(i))

                        Dim MaterialNo As String = CLIDs.Tables(0).Rows(i)("MaterialNo").ToString
                        DR = dtItem.Select("MaterialNo = '" & MaterialNo & "'")
                        If DR.Length = 0 Then
                            j = j + 1
                            myDR = dtItem.NewRow
                            myDR("MaterialNo") = MaterialNo
                            myDR("PNSeq") = j
                            dtItem.Rows.Add(myDR)
                            SeqFlag = True
                        End If
                    End If
                Next
                If SeqFlag = False Then Exit Function


                Dim SqlStr As String
                Dim ds As New DataSet()

                Try
                    CLIDs = New DataSet
                    CLIDs.Tables.Add(dtCLID)

                    CLIDs.DataSetName = "CLIDs"
                    CLIDs.Tables(0).TableName = "dtCLID"

                    SqlStr = String.Format("exec dbo.sp_RECReadLabels  N'{0}' ", DStoXML(CLIDs))
                    ds = da.ExecuteDataSet(SqlStr, "dtCLID")
                Catch ex As Exception
                    ErrorLogging("Receiving-Call:sp_RECReadLabels", "", ex.Message & ex.Source, "E")
                    ds = Nothing
                End Try

                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    PrintRECLabels = False
                    Exit Function
                End If

                'Default set to print out Barcode for Material, if this Vendor is on AML Verification lists, then do not print out Barcode for Material
                Dim IPNFlag As String = "Y"
                Dim tmpVendor As String = ""
                Dim POVendor As String = ds.Tables(0).Rows(0)("VendorID").ToString
                SqlStr = String.Format("Select ProcessID as VendorNo from T_SysLOV with (nolock) where Name = 'Vendors for CLID Verification' and ProcessID = '" & POVendor & "'")
                tmpVendor = Convert.ToString(da.ExecuteScalar(SqlStr))
                If tmpVendor <> "" Then IPNFlag = "N"


                Dim lblPrint As String
                Dim TmpDate As Date
                Dim lblData As MatLabel = New MatLabel

                '0  CLID, OrgCode, MaterialNo, MaterialRevision, QtyBaseUOM, BaseUOM, LotNo, DateCode, ExpDate, RecDate, RecDocNo, RoHS, StockType, '12
                '13 MatSuffix1, MatSuffix2, MatSuffix3, StatusCode, AddlData, Manufacturer, ManufacturerPN, Stemp, MSL, RTLot, MCPosition, PredefinedSubInv, ItemText, COO  '26
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Dim CLID As String = ds.Tables(0).Rows(i)(0).ToString
                    If Microsoft.VisualBasic.Mid(CLID, 3, 1) = "B" Then
                        lblData.LabelType = "Box Label"
                    Else
                        lblData.LabelType = "Material Label"
                    End If

                    If Not ds.Tables(0).Rows(i)(0) Is DBNull.Value Then lblData.LabelID = ds.Tables(0).Rows(i)(0)
                    If Not ds.Tables(0).Rows(i)(1) Is DBNull.Value Then lblData.OrgCode = ds.Tables(0).Rows(i)(1)
                    If Not ds.Tables(0).Rows(i)(2) Is DBNull.Value Then lblData.Material = ds.Tables(0).Rows(i)(2)
                    If Not ds.Tables(0).Rows(i)(3) Is DBNull.Value Then lblData.Rev = ds.Tables(0).Rows(i)(3)
                    If Not ds.Tables(0).Rows(i)(4) Is DBNull.Value Then lblData.Qty = ds.Tables(0).Rows(i)(4)
                    If Not ds.Tables(0).Rows(i)(5) Is DBNull.Value Then lblData.UOM = ds.Tables(0).Rows(i)(5)
                    If Not ds.Tables(0).Rows(i)(6) Is DBNull.Value Then lblData.LotNo = ds.Tables(0).Rows(i)(6)
                    If Not ds.Tables(0).Rows(i)(7) Is DBNull.Value Then lblData.DCode = ds.Tables(0).Rows(i)(7)

                    lblData.Qty = Format(ds.Tables(0).Rows(i)(4), "#0.#####")

                    If Not ds.Tables(0).Rows(i)(8) Is DBNull.Value Then
                        TmpDate = ds.Tables(0).Rows(i)(8)
                        lblData.ExpDate = TmpDate.ToString("MM/dd/yyyy")
                    End If

                    If Not ds.Tables(0).Rows(i)(9) Is DBNull.Value Then
                        TmpDate = ds.Tables(0).Rows(i)(9)
                        lblData.RecDate = TmpDate.ToString("MM/dd/yyyy")
                    End If

                    Dim MaterialNo As String = ds.Tables(0).Rows(i)("MaterialNo").ToString
                    DR = dtItem.Select("MaterialNo = '" & MaterialNo & "'")
                    If DR.Length > 0 Then lblData.PNSeq = DR(0)("PNSeq").ToString

                    'Default set to print out Barcode for Material, if this Vendor is on AML Verification lists, then do not print out Barcode for Material
                    lblData.IPN = lblData.Material
                    If IPNFlag = "N" Then lblData.IPN = ""

                    If Not ds.Tables(0).Rows(i)(10) Is DBNull.Value Then lblData.RT = ds.Tables(0).Rows(i)(10)
                    If Not ds.Tables(0).Rows(i)(11) Is DBNull.Value Then lblData.RoHS = ds.Tables(0).Rows(i)(11)
                    If Not ds.Tables(0).Rows(i)(12) Is DBNull.Value Then lblData.InspFlag = ds.Tables(0).Rows(i)(12)
                    If Not ds.Tables(0).Rows(i)(18) Is DBNull.Value Then lblData.Mfr = ds.Tables(0).Rows(i)(18)
                    If Not ds.Tables(0).Rows(i)(19) Is DBNull.Value Then lblData.MfrPN = ds.Tables(0).Rows(i)(19)
                    If Not ds.Tables(0).Rows(i)(20) Is DBNull.Value Then lblData.Stemp = ds.Tables(0).Rows(i)(20)
                    If Not ds.Tables(0).Rows(i)(21) Is DBNull.Value Then lblData.MSL = ds.Tables(0).Rows(i)(21)

                    'Check if RTLot is different from RT, if yes, then Combine RT / RTLot into field lblData.RT
                    If ds.Tables(0).Rows(i)(22) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(i)(22)) = "" Then
                    Else
                        Dim RT As String = lblData.RT
                        Dim RTLot As String = Trim(ds.Tables(0).Rows(i)(22))
                        'If RTLot <> RT Then lblData.RT = lblData.RT & " / " & RTLot
                        If RTLot <> RT Then                                                 'RTLot generated by RT-ExpDate (yyMMdd)
                            If RTLot.Contains("-") Then
                                lblData.RT = RTLot
                            Else
                                lblData.RT = lblData.RT & " / " & RTLot
                            End If
                        End If
                    End If

                    'Print MCPosition on MaterialLabel requested by ZS   --- 06/24/2014
                    If Not ds.Tables(0).Rows(i)(23) Is DBNull.Value Then lblData.MCPosition = ds.Tables(0).Rows(i)(23)

                    'Print PredefinedSubInv on MaterialLabel requested by ZS only  --- 03/27/2015
                    If Not ds.Tables(0).Rows(i)(24) Is DBNull.Value Then lblData.PreSubInv = ds.Tables(0).Rows(i)(24)

                    'Print ItemText on MaterialLabel requested by LD Huawei project  --- 12/27/2016
                    If Not ds.Tables(0).Rows(i)(25) Is DBNull.Value Then lblData.ItemText = ds.Tables(0).Rows(i)(25)

                    'Print COO on MaterialLabel requested by ZS only  --- 12/27/2016
                    If Not ds.Tables(0).Rows(i)(26) Is DBNull.Value Then lblData.COO = ds.Tables(0).Rows(i)(26)

                    lblData.Safety = ""
                    lblData.ESD = BlankESD
                    If Not ds.Tables(0).Rows(i)(17) Is DBNull.Value Then
                        Dim AddlData As String = ds.Tables(0).Rows(i)(17).ToString.ToUpper
                        If InStr(AddlData, "S") > 0 Then lblData.Safety = "Safety"
                        If InStr(AddlData, "E") > 0 Then lblData.ESD = ESDImage
                    End If

                    Dim flag As Integer = 0
                    If ds.Tables(0).Rows(i)(13) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(i)(13)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & ds.Tables(0).Rows(i)(13) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If
                    If ds.Tables(0).Rows(i)(14) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(i)(14)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & ds.Tables(0).Rows(i)(14) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If
                    If ds.Tables(0).Rows(i)(15) Is DBNull.Value OrElse Trim(ds.Tables(0).Rows(i)(15)) = "" Then
                    Else
                        lblData.MatSuffix = "[" & ds.Tables(0).Rows(i)(15) & "]"
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                        flag = 1
                    End If

                    If flag = 0 Then
                        lblPrint = PrintLabel(LabelPrinter, lblData)
                    End If
                    If lblPrint <> "True" Then PrintRECLabels = False
                    Sleep(5)
                Next

            Catch ex As Exception
                ErrorLogging("Receiving-PrintRECLabels", "", ex.Message & ex.Source, "E")
                PrintRECLabels = False
            End Try
        End Using

    End Function

    Public Function PrintRTSlip(ByVal LoginData As ERPLogin, ByVal RTItems As DataSet, ByVal LabelPrinter As String) As Boolean
        Using da As DataAccess = GetDataAccess()
            PrintRTSlip = True
            Dim RTNo As String = RTItems.Tables(0).Rows(0)("RTNo").ToString

            Try
                Dim mySqlstr As String
                mySqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC014'")
                Dim RTprt As String = Convert.ToString(da.ExecuteScalar(mySqlstr)).ToUpper
                If RTprt = "YES" Then
                    'Sort RT by Material Ascending before written to eTrace table T_LabelPrint           -- 11/21/2018
                    'Please note this is only for ZS site,  LD/PH will NOT sort Material here
                    Dim dtPrint As DataTable = New DataTable
                    Dim SortColName As String = RTItems.Tables(0).Columns(2).ColumnName                    'MaterialNo
                    SortColName = SortColName & " ASC"

                    RTItems.Tables(0).DefaultView.Sort = SortColName
                    dtPrint = RTItems.Tables(0).DefaultView.ToTable()

                    RTItems = New DataSet
                    RTItems.Tables.Add(dtPrint)
                End If


                Dim RTData As New DataSet
                RTData = GetRTData(LoginData, RTItems)

                If RTData Is Nothing OrElse RTData.Tables.Count = 0 OrElse RTData.Tables(0).Rows.Count = 0 Then
                    PrintRTSlip = False
                    Exit Function
                End If

                Dim arryFile() As String
                arryFile = Split(RTLabelFile, "\")
                Dim LabelFile As String = arryFile(UBound(arryFile))

                Dim i, aa As Integer
                For i = 0 To RTData.Tables(0).Rows.Count - 1
                    Dim RTContent As String = RTData.Tables(0).Rows(i)(0).ToString

                    'Filter Special Characters for RTData
                    RTContent = FilterSpecial(RTContent)

                    Dim sqlstr As String
                    sqlstr = String.Format("exec sp_InsertLabelPrint '{0}','{1}','{2}'", LabelFile, LabelPrinter, SQLString(RTContent))
                    da.ExecuteScalar(sqlstr)

                    'sqlstr = String.Format("INSERT INTO T_LabelPrint (LabelSeqNo,LabelFile,LabelPrinter,Content,CreatedOn,StatusCode) values (newid(),'{0}','{1}','{2}',getdate(),1)", LabelFile, LabelPrinter, RTContent)
                    'aa = CInt(da.ExecuteNonQuery(sqlstr))
                    'If aa < 1 Then PrintRTSlip = False
                Next
            Catch ex As Exception
                ErrorLogging("Receiving-PrintRTSlip", LoginData.User.ToUpper, "RT: " & RTNo & ex.Message & ex.Source, "E")
                PrintRTSlip = False
            End Try
        End Using

    End Function

    Private Function GetRTData(ByVal LoginData As ERPLogin, ByVal RTItems As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            GetRTData = New DataSet
            Dim RTNo As String = RTItems.Tables(0).Rows(0)("RTNo").ToString

            Try
                Dim i As Integer
                Dim Sqlstr As String
                For i = 0 To RTItems.Tables(0).Rows.Count - 1
                    Dim TransactionID As String = RTItems.Tables(0).Rows(i)("TransactionID").ToString
                    Sqlstr = String.Format("Select RTContent from T_RTSlip with (nolock) where TransactionID = '{0}' ", TransactionID)

                    Dim ds As DataSet = New DataSet()
                    ds = da.ExecuteDataSet(Sqlstr, "RTTable")
                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                        ErrorLogging("Receiving-GetRTData1", LoginData.User.ToUpper, " Couldn't find RTContent for RT: " & RTNo & " with TransactionID " & TransactionID & " in table T_RTSlip", "I")
                        Continue For
                    End If


                    Dim NewRT As String = ""
                    Dim UpdateFlag As Boolean = False
                    Dim RTContent As String = ds.Tables(0).Rows(0)(0).ToString

                    'Check if the RTContent contained new column MCPosition, if no, added this column to match the new RTSlip format
                    If Not RTContent.Contains("MCPosition") Then
                        UpdateFlag = True
                        NewRT = Replace(RTContent, "^StorageType^", "^MCPosition^^StorageType^")
                        RTContent = NewRT
                    End If

                    'Check if the RTContent contained new column POType, if no, added this column to match the new RTSlip format
                    If Not RTContent.Contains("POType") Then
                        UpdateFlag = True
                        NewRT = Replace(RTContent, "^Buyer^", "^POType^^Buyer^")
                    End If

                    'Check if the RTContent contained new column VMIFlag, if no, added this column to match the new RTSlip format
                    If Not RTContent.Contains("VMIFlag") Then
                        UpdateFlag = True
                        NewRT = Replace(RTContent, "^POType^", "^VMIFlag^^POType^")
                    End If

                    'Check if the RTContent contained new column DateCode/LotNo, if no, added this column to match the new RTSlip format
                    If Not RTContent.Contains("DateCode") Then
                        UpdateFlag = True
                        NewRT = RTContent & "^DateCode^^LotNo^"
                    End If

                    'Check if the RTContent contained new column ESD/ShelfLife, if no, added these two columns to match the new RTSlip format      -- 05/04/2017
                    If Not RTContent.Contains("ShelfLife") Then
                        UpdateFlag = True
                        NewRT = RTContent & "^ESD^^ShelfLife^"
                    End If

                    'Check if the RTContent contained new column HisInsRT, if no, added this column to match the new RTSlip format        -- 02/21/2019
                    If Not RTContent.Contains("HisInsRT") Then
                        UpdateFlag = True
                        NewRT = RTContent & "^HisInsRT^"
                    End If

                    'Check if the RTContent contained new column SampleLoc, if no, added this column to match the new RTSlip format      -- 02/21/2019
                    If Not RTContent.Contains("SampleLoc") Then
                        UpdateFlag = True
                        NewRT = RTContent & "^SampleLoc^"
                    End If

                    'Update RTContent in table T_RTSlip to add these columns for future printing use
                    If UpdateFlag = True Then
                        ds.Tables(0).Rows(0)(0) = NewRT

                        'Filter Special Characters for NewRT
                        NewRT = FilterSpecial(NewRT)

                        Sqlstr = String.Format("UPDATE T_RTSlip set RTContent='{0}' where TransactionID = '{1}' ", NewRT, TransactionID)
                        da.ExecuteNonQuery(Sqlstr)
                    End If


                    If GetRTData Is Nothing OrElse GetRTData.Tables.Count = 0 Then
                        GetRTData = ds
                    Else
                        GetRTData.Tables(0).Merge(ds.Tables(0))
                    End If
                Next

                Return GetRTData

            Catch ex As Exception
                ErrorLogging("Receiving-GetRTData", LoginData.User.ToUpper, "RT: " & RTNo & ex.Message & ex.Source, "E")
                GetRTData = Nothing
            End Try
        End Using

    End Function

    Public Function Get_HeaderInterfaceID() As Integer
        Using da As DataAccess = GetDataAccess()
            Get_HeaderInterfaceID = 0

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()
            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.get_rcv_header_id"
                OC.Parameters.Add("o_head_inte_id", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Get_HeaderInterfaceID = CInt(OC.Parameters("o_head_inte_id").Value)
                OC.Connection.Close()
                'OC.Dispose()

                Return Get_HeaderInterfaceID

            Catch oe As Exception
                ErrorLogging("Receiving-Get_HeaderInterfaceID", "", oe.Message & oe.Source, "E")
                Return 0
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_HeaderGroupID() As RcvID
        Using da As DataAccess = GetDataAccess()
            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Dim myRcvID As RcvID = New RcvID
            myRcvID.HeaderID = 0
            myRcvID.GroupID = 0

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.get_rcv_head_group_id"
                OC.Parameters.Add("o_head_inte_id", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_group_id", OracleType.VarChar, 240).Direction = ParameterDirection.Output
                OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                myRcvID.HeaderID = CInt(OC.Parameters("o_head_inte_id").Value)
                myRcvID.GroupID = CInt(OC.Parameters("o_group_id").Value)

                OC.Connection.Close()

                Return myRcvID

            Catch oe As Exception
                ErrorLogging("Receiving-Get_HeaderGroupID", "", oe.Message & oe.Source, "E")
                Return myRcvID
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_PO_Detail(ByVal OrgCode As String, ByVal GRType As Integer, ByVal OrderData As POData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()    'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num

            Dim OrderNo As String = OrderData.PONo
            If OrderData.ReleaseNo > 0 Then OrderNo = OrderData.PONo & "-" & OrderData.ReleaseNo

            Try
                ds.Tables.Add("po_header")
                ds.Tables.Add("po_line")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                If OrderData.ReleaseNo > 0 Then
                    If GRType = 2 Then
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_bpa_osp_header_data"      'Get Release BPA for OSP PO
                    Else
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_bpa_header_data"          'Get Release BPA for Kanban PO
                    End If
                    oda.SelectCommand.Parameters.Add("p_release_num", OracleType.VarChar, 240).Value = CStr(OrderData.ReleaseNo)
                Else
                    If GRType = 2 Then
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_osp_header_data"          'Get Standard OSP PO
                    Else
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_header_data"              'Get Standard PO
                    End If
                End If
                oda.SelectCommand.Parameters.Add("o_po_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_po_no", OracleType.VarChar, 240).Value = CStr(OrderData.PONo)
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "po_header")

                If OrderData.ReleaseNo > 0 Then
                    If GRType = 2 Then
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_bpa_osp_line_data"        'Get Release BPA Line for OSP PO
                    Else
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_bpa_line_data"            'Get Release BPA Line for Kanban PO
                    End If
                Else
                    If GRType = 2 Then
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_osp_line_data"            'Get Standard OSP PO Line
                    Else
                        oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_line_data"                'Get Standard PO Line
                    End If
                End If

                If OrderData.LineNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_line_num", OracleType.Number).Value = OrderData.LineNo
                End If
                If OrderData.ShipmentNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.Number).Value = OrderData.ShipmentNo
                End If
                oda.Fill(ds, "po_line")
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_PO_Detail", "", "PONo: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Get_Batch_PO(ByVal OrgCode As String, ByVal BatchPO As DataTable) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()               'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num

            Try
                Dim i As Integer
                Dim POLists As String = ""

                For i = 0 To BatchPO.Rows.Count - 1
                    Dim OrderData As POData
                    Dim OrderNo, OrderItem, StrPO As String

                    OrderNo = BatchPO.Rows(i)("PONo").ToString
                    OrderItem = BatchPO.Rows(i)("POLine").ToString
                    OrderData = Split_POData(OrderNo, OrderItem)

                    StrPO = OrderData.PONo & "," & OrderData.ReleaseNo & "," & OrderData.LineNo & "," & OrderData.ShipmentNo & ";"
                    If i = 0 Then
                        POLists = StrPO
                    Else
                        POLists = POLists & StrPO
                    End If
                Next

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_batch_data"              'Get Standard PO

                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_po_data", OracleType.VarChar, 30000).Value = POLists
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                'oda.Fill(ds, "po_header")
                'oda.Fill(ds, "po_line")
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_Batch_PO", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_Batch_EJITPO(ByVal OrgCode As String, ByVal BatchPO As DataTable) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                BatchPO.TableName = "dtItem"
                BatchPO.Columns(0).ColumnName = "pono"
                BatchPO.Columns(1).ColumnName = "poline"

                Dim dsItem As New DataSet
                dsItem.DataSetName = "dsItem"
                dsItem.Tables.Add(BatchPO)

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_batch_ejitkb_po_data"              'Get batch ejit PO
                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Parameters.Add("p_xml", OracleType.Clob).Value = DStoXML(dsItem)
                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_Batch_EJITPO", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_Batch_ASN(ByVal OrgCode As String, ByVal BatchPO As DataTable) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                BatchPO.TableName = "dtItem"
                BatchPO.Columns(0).ColumnName = "asn_number"
                BatchPO.Columns(1).ColumnName = "asn_line_number"

                Dim dsItem As New DataSet
                dsItem.DataSetName = "dsItem"
                dsItem.Tables.Add(BatchPO)

                'o_po_header         OUT cur_po,
                'o_po_line           OUT cur_po,
                'p_org_code          IN VARCHAR2,
                'p_xml               IN CLOB
                '<dsItem>
                '  <dtItem>
                '    <asn_number>ASBN-ZH_02</asn_number>
                '    <asn_line_number>2</asn_line_number>
                '  </dtItem>
                '  <dtItem>
                '    <asn_number>ASBN-ZH_02</asn_number>
                '    <asn_line_number>3</asn_line_number>
                '  </dtItem>
                '</dsItem>' 

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_batch_asn_po_data"              'Get batch ASN for Normal PO / ejit po
                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Parameters.Add("p_xml", OracleType.Clob).Value = DStoXML(dsItem)
                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_Batch_ASN", "", oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_PO_Kanban(ByVal OrgCode As String, ByVal OrderData As POData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Dim OrderNo As String = OrderData.PONo
            If OrderData.ReleaseNo > 0 Then OrderNo = OrderData.PONo & "-" & OrderData.ReleaseNo

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_kanban_data"              
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_po_kanban_data_v2"              'New PKG to get KanBan PO

                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_po_no", OracleType.VarChar, 240).Value = OrderData.PONo
                oda.SelectCommand.Parameters.Add("p_release_num", OracleType.VarChar, 240).Value = CStr(OrderData.ReleaseNo)
                oda.SelectCommand.Parameters.Add("p_line_num", OracleType.Number).Value = OrderData.LineNo
                oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.Number).Value = OrderData.ShipmentNo
                oda.SelectCommand.Parameters.Add("p_distribution_num", OracleType.Number).Value = OrderData.Distribution
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_PO_Kanban", "", "PONo: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Get_IR_ISO(ByVal OrgCode As String, ByVal OrderData As POData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Dim OrderNo As String = OrderData.PONo
            If OrderData.ReleaseNo > 0 Then OrderNo = OrderData.PONo & "-" & OrderData.ReleaseNo

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_ir_so_data"              'Get IR / ISO

                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.VarChar, 50).Value = OrderData.PONo
                oda.SelectCommand.Parameters.Add("p_line_num", OracleType.Int32).Value = OrderData.LineNo
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode

                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_IR_ISO", "", "ShipmentNo: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function Get_PO_Indirect(ByVal OrgCode As String, ByVal OrderData As POData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Dim OrderNo As String = OrderData.PONo
            If OrderData.ReleaseNo > 0 Then OrderNo = OrderData.PONo & "-" & OrderData.ReleaseNo

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_indirect_po_data"              'Get Indirect PO Data

                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Parameters.Add("p_po_num", OracleType.Number).Value = OrderData.PONo

                If OrderData.LineNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_po_line_num", OracleType.Number).Value = OrderData.LineNo
                End If
                If OrderData.ShipmentNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_po_shipment_num", OracleType.Number).Value = OrderData.ShipmentNo
                End If

                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_PO_Indirect", "", "PONo: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try

        End Using

    End Function

    Public Function Get_ASN_PO(ByVal OrgCode As String, ByVal OrderData As POData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Dim OrderNo As String = OrderData.PONo
            If OrderData.ReleaseNo > 0 Then OrderNo = OrderData.PONo & "-" & OrderData.ReleaseNo

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_asn_po_data"              'Get ASN PO

                oda.SelectCommand.Parameters.Add("o_po_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_po_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Parameters.Add("p_asn_shipment_num", OracleType.VarChar, 50).Value = OrderData.PONo

                If OrderData.LineNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_po_line_num", OracleType.Int32).Value = OrderData.LineNo
                End If
                If OrderData.ShipmentNo > 0 Then
                    oda.SelectCommand.Parameters.Add("p_po_shipment_num", OracleType.VarChar, 50).Value = OrderData.ShipmentNo
                End If

                oda.SelectCommand.Connection.Open()

                oda.Fill(ds)
                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_ASN_PO", "", "ASN: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Get_RT_Data(ByVal OrgCode As String, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet()
            Dim oda As OracleDataAdapter = da.Oda_Sele()    'p_org_code	p_po_no	p_release_num	p_line_num	p_shipment_num

            Try

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_po_detail_pkg.get_rt_data"              'Get Receipt Number
                oda.SelectCommand.Parameters.Add("o_rt_header", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("o_rt_line", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_rt_num", OracleType.VarChar, 100).Value = OrderNo
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = OrgCode
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()

                If ds Is Nothing OrElse ds.Tables.Count < 2 Then
                    Return Nothing
                    Exit Function
                End If

                ds.Tables(0).TableName = "po_header"
                ds.Tables(1).TableName = "po_line"

                Return ds

            Catch oe As Exception
                ErrorLogging("Receiving-Get_RT_Data", "", "ReceiptNumber: " & OrderNo & ", " & oe.Message & oe.Source, "E")
                Return Nothing
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Post_Receive(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet, ByVal p_batch_id As Integer, ByVal Header As GRHeaderStructure) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Dim GRResult As CreateGRResponse = New CreateGRResponse

            Dim ErrorTable As DataTable
            Dim ErrorRow As Data.DataRow
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ColumnName", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables("po_line").Rows(0)("PO_NUMBER"), 4, 2)
            If Header.Type = 4 Then POStr = "IR"

            Dim oda_h As OracleDataAdapter = da.Oda_Insert()
            Dim oda_l As OracleDataAdapter = da.Oda_Insert()

            Try
                oda_h.InsertCommand.CommandType = CommandType.StoredProcedure

                If Header.Type = 4 Then                   'IRISO
                    oda_h.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_header_irso"
                    oda_h.InsertCommand.Parameters.Add("p_shipment_num", OracleType.VarChar, 240)
                    oda_h.InsertCommand.Parameters.Add("p_ship_to_organization_id", OracleType.Int32)

                    oda_h.InsertCommand.Parameters("p_shipment_num").SourceColumn = "shipment_num"
                    oda_h.InsertCommand.Parameters("p_ship_to_organization_id").SourceColumn = "ship_to_organization_id"
                ElseIf Header.Type = 6 Or Header.Type = 7 Then                   'ASN
                    oda_h.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_header_asnpo"
                    oda_h.InsertCommand.Parameters.Add("p_shipment_header_id", OracleType.VarChar, 240)
                    oda_h.InsertCommand.Parameters.Add("p_shipment_num", OracleType.VarChar, 240)
                    oda_h.InsertCommand.Parameters.Add("p_ship_to_organization_id", OracleType.Int32)
                    oda_h.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)

                    oda_h.InsertCommand.Parameters("p_shipment_header_id").SourceColumn = "shipment_header_id"
                    oda_h.InsertCommand.Parameters("p_shipment_num").SourceColumn = "shipment_num"
                    oda_h.InsertCommand.Parameters("p_ship_to_organization_id").SourceColumn = "ship_to_organization_id"
                    oda_h.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                    oda_h.InsertCommand.Parameters.Add("o_rt_num", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                Else
                    oda_h.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_header_interface"          ' 1 /3 
                    oda_h.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)
                    oda_h.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                End If

                oda_h.InsertCommand.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
                oda_h.InsertCommand.Parameters.Add("p_header_interface_id", OracleType.Int32)
                oda_h.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                oda_h.InsertCommand.Parameters.Add("p_headers_comments", OracleType.VarChar, 240)
                oda_h.InsertCommand.Parameters.Add("p_waybill_airbill_num", OracleType.VarChar, 240)
                oda_h.InsertCommand.Parameters.Add("p_bill_of_lading", OracleType.VarChar, 240)
                oda_h.InsertCommand.Parameters.Add("p_packing_slip", OracleType.VarChar, 240)
                oda_h.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda_h.InsertCommand.Parameters("p_header_interface_id").SourceColumn = "header_interface_id"
                oda_h.InsertCommand.Parameters("p_user_id").SourceColumn = "user_id"
                oda_h.InsertCommand.Parameters("p_headers_comments").SourceColumn = "comments"
                oda_h.InsertCommand.Parameters("p_waybill_airbill_num").SourceColumn = "waybill_airbill_num"
                oda_h.InsertCommand.Parameters("p_bill_of_lading").SourceColumn = "bill_of_lading"
                oda_h.InsertCommand.Parameters("p_packing_slip").SourceColumn = "packing_slip"
                oda_h.InsertCommand.Parameters("o_succ_flag").SourceColumn = "succ_flag"

                oda_h.InsertCommand.Connection.Open()
                oda_h.Update(p_ds.Tables("po_header"))
                oda_h.InsertCommand.Connection.Close()
                'oda_h.Dispose()

                If Header.Type = 7 Then
                    Dim RTNo As String = FixNull(oda_h.InsertCommand.Parameters("o_rt_num").Value)
                    Dim j As Integer
                    For j = 0 To p_ds.Tables("po_line").Rows.Count - 1
                        Dim LotNumber As String
                        If p_ds.Tables("po_line").Rows(j)("lot_expiration_date") Is DBNull.Value Then
                            LotNumber = Format(Header.PostDate, "yyMMdd")
                        Else
                            LotNumber = Format(p_ds.Tables("po_line").Rows(j)("lot_expiration_date"), "yyMMdd")
                        End If
                        If Header.POCurrency = "CNY" Then LotNumber = LotNumber & "CN"
                        p_ds.Tables("po_line").Rows(j)("lot_num") = RTNo & "-" & LotNumber

                        If p_ds.Tables("po_line").Rows(j).RowState = DataRowState.Unchanged Then
                            p_ds.Tables("po_line").Rows(j).SetAdded()
                        End If
                    Next
                End If

                oda_l.InsertCommand.CommandType = CommandType.StoredProcedure

                If Header.Type = 1 Then
                    oda_l.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_only_rec"

                ElseIf Header.Type > 1 Then

                    If Header.Type = 3 Then
                        oda_l.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_dire"
                        oda_l.InsertCommand.Parameters.Add("p_locator", OracleType.VarChar, 240)             'Add Locator
                        oda_l.InsertCommand.Parameters("p_locator").SourceColumn = "destination_locator"

                    ElseIf Header.Type = 4 Then
                        oda_l.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_irso"
                        oda_l.InsertCommand.Parameters.Add("p_shipment_num", OracleType.VarChar, 240)
                        oda_l.InsertCommand.Parameters.Add("p_shipment_header_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_shipment_line_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_requisition_head_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_requisition_line_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_req_distribution_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_ship_to_location_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_ej_destination_locator", OracleType.VarChar, 240)
                        oda_l.InsertCommand.Parameters.Add("p_lot_num", OracleType.VarChar, 240)

                        oda_l.InsertCommand.Parameters("p_shipment_num").SourceColumn = "po_number"
                        oda_l.InsertCommand.Parameters("p_shipment_header_id").SourceColumn = "shipment_header_id"
                        oda_l.InsertCommand.Parameters("p_shipment_line_id").SourceColumn = "shipment_line_id"
                        oda_l.InsertCommand.Parameters("p_requisition_head_id").SourceColumn = "requisition_head_id"
                        oda_l.InsertCommand.Parameters("p_requisition_line_id").SourceColumn = "requisition_line_id"
                        oda_l.InsertCommand.Parameters("p_req_distribution_id").SourceColumn = "po_distribution_id"
                        oda_l.InsertCommand.Parameters("p_ship_to_location_id").SourceColumn = "ship_to_location_id"
                        oda_l.InsertCommand.Parameters("p_ej_destination_locator").SourceColumn = "destination_locator"
                        oda_l.InsertCommand.Parameters("p_lot_num").SourceColumn = "rt_lot_number"

                    ElseIf Header.Type = 5 Then                   'Indirect PO
                        oda_l.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_indirectpo"
                        oda_l.InsertCommand.Parameters.Add("p_vendor_site_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_po_unit_price", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_item_description", OracleType.VarChar, 250)

                        oda_l.InsertCommand.Parameters("p_vendor_site_id").SourceColumn = "vendor_site_id"
                        oda_l.InsertCommand.Parameters("p_po_unit_price").SourceColumn = "po_unit_price"
                        oda_l.InsertCommand.Parameters("p_item_description").SourceColumn = "item_description"

                    ElseIf Header.Type = 6 Or Header.Type = 7 Then                  'ASN PO / eJit PO
                        oda_l.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_asnpo"
                        oda_l.InsertCommand.Parameters.Add("p_shipment_header_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_shipment_line_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_shipment_num", OracleType.VarChar, 240)
                        oda_l.InsertCommand.Parameters.Add("p_ship_to_location_id", OracleType.Int32)

                        oda_l.InsertCommand.Parameters("p_shipment_header_id").SourceColumn = "shipment_header_id"
                        oda_l.InsertCommand.Parameters("p_shipment_line_id").SourceColumn = "shipment_line_id"
                        oda_l.InsertCommand.Parameters("p_shipment_num").SourceColumn = "asn_shipment_num"
                        oda_l.InsertCommand.Parameters("p_ship_to_location_id").SourceColumn = "ship_to_location_id"

                        If Header.Type = 7 Then
                            oda_l.InsertCommand.Parameters.Add("p_lot_num", OracleType.VarChar, 240)
                            oda_l.InsertCommand.Parameters("p_lot_num").SourceColumn = "lot_num"
                        End If
                    End If

                    oda_l.InsertCommand.Parameters.Add("p_primary_quantity", OracleType.[Double])
                    oda_l.InsertCommand.Parameters.Add("p_primary_unit_of_measure", OracleType.VarChar, 50)
                    oda_l.InsertCommand.Parameters.Add("p_deliver_to_location_id", OracleType.Int32)
                    oda_l.InsertCommand.Parameters("p_primary_quantity").SourceColumn = "REC_PRIMARY_QTY"
                    oda_l.InsertCommand.Parameters("p_primary_unit_of_measure").SourceColumn = "base_uom"  'base_uom_code
                    oda_l.InsertCommand.Parameters("p_deliver_to_location_id").SourceColumn = "deliver_to_location_id"

                    If Header.Type = 3 Or Header.Type = 4 Or Header.Type = 6 Or Header.Type = 7 Then
                        oda_l.InsertCommand.Parameters.Add("p_deliver_to_person_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters("p_deliver_to_person_id").SourceColumn = "deliver_to_person_id"

                        If Header.Type = 3 Or Header.Type = 4 Or Header.Type = 7 Then
                            oda_l.InsertCommand.Parameters.Add("p_subinventory", OracleType.VarChar, 240)
                            oda_l.InsertCommand.Parameters.Add("p_locator_id", OracleType.Int32)
                            oda_l.InsertCommand.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
                            oda_l.InsertCommand.Parameters("p_subinventory").SourceColumn = "destination_subinventory"
                            oda_l.InsertCommand.Parameters("p_locator_id").SourceColumn = "locator_id"
                            oda_l.InsertCommand.Parameters("p_lot_expiration_date").SourceColumn = "lot_expiration_date"
                        End If
                    End If
                    'If Header.Type <> 5 Then
                    '    oda_l.InsertCommand.Parameters.Add("p_deliver_to_person_id", OracleType.Int32)
                    '    oda_l.InsertCommand.Parameters("p_deliver_to_person_id").SourceColumn = "deliver_to_person_id"

                    '    If Header.Type <> 6 Then
                    '        oda_l.InsertCommand.Parameters.Add("p_subinventory", OracleType.VarChar, 240)
                    '        oda_l.InsertCommand.Parameters.Add("p_locator_id", OracleType.Int32)
                    '        oda_l.InsertCommand.Parameters.Add("p_lot_expiration_date", OracleType.DateTime)
                    '        oda_l.InsertCommand.Parameters("p_subinventory").SourceColumn = "destination_subinventory"
                    '        oda_l.InsertCommand.Parameters("p_locator_id").SourceColumn = "locator_id"
                    '        oda_l.InsertCommand.Parameters("p_lot_expiration_date").SourceColumn = "lot_expiration_date"
                    '    End If
                    'End If
                End If

                'If Header.Type <> 4 Then
                '    oda_l.InsertCommand.Parameters.Add("p_po_header_id", OracleType.Int32)
                '    oda_l.InsertCommand.Parameters.Add("p_po_line_id", OracleType.Int32)
                '    oda_l.InsertCommand.Parameters.Add("p_po_line_location_id", OracleType.Int32)
                '    oda_l.InsertCommand.Parameters("p_po_header_id").SourceColumn = "po_header_id"
                '    oda_l.InsertCommand.Parameters("p_po_line_id").SourceColumn = "po_line_id"
                '    oda_l.InsertCommand.Parameters("p_po_line_location_id").SourceColumn = "po_line_location_id"

                '    If Header.Type <> 6 And Header.Type <> 7 Then
                '        oda_l.InsertCommand.Parameters.Add("p_po_release_id", OracleType.Int32)        'p_po_release_id
                '        oda_l.InsertCommand.Parameters.Add("p_line_num", OracleType.Int32)
                '        oda_l.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)
                '        oda_l.InsertCommand.Parameters.Add("p_po_distribution_id", OracleType.Int32)
                '        oda_l.InsertCommand.Parameters("p_po_release_id").SourceColumn = "po_release_id"
                '        oda_l.InsertCommand.Parameters("p_line_num").SourceColumn = "line_num"
                '        oda_l.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                '        oda_l.InsertCommand.Parameters("p_po_distribution_id").SourceColumn = "po_distribution_id"
                '    End If
                'End If

                If Header.Type = 1 Or Header.Type = 3 Or Header.Type >= 5 Then
                    oda_l.InsertCommand.Parameters.Add("p_po_header_id", OracleType.Int32)
                    oda_l.InsertCommand.Parameters.Add("p_po_line_id", OracleType.Int32)
                    oda_l.InsertCommand.Parameters.Add("p_po_line_location_id", OracleType.Int32)
                    oda_l.InsertCommand.Parameters("p_po_header_id").SourceColumn = "po_header_id"
                    oda_l.InsertCommand.Parameters("p_po_line_id").SourceColumn = "po_line_id"
                    oda_l.InsertCommand.Parameters("p_po_line_location_id").SourceColumn = "po_line_location_id"

                    If Header.Type = 1 Or Header.Type = 3 Or Header.Type = 5 Then
                        oda_l.InsertCommand.Parameters.Add("p_po_release_id", OracleType.Int32)        'p_po_release_id
                        oda_l.InsertCommand.Parameters.Add("p_line_num", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)
                        oda_l.InsertCommand.Parameters.Add("p_po_distribution_id", OracleType.Int32)

                        oda_l.InsertCommand.Parameters("p_po_release_id").SourceColumn = "po_release_id"
                        oda_l.InsertCommand.Parameters("p_line_num").SourceColumn = "line_num"
                        oda_l.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                        oda_l.InsertCommand.Parameters("p_po_distribution_id").SourceColumn = "po_distribution_id"
                    End If
                End If


                oda_l.InsertCommand.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
                oda_l.InsertCommand.Parameters.Add("p_header_interface_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_item_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_item_category_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_item_revision", OracleType.VarChar, 50)
                oda_l.InsertCommand.Parameters.Add("p_quantity", OracleType.[Double])
                oda_l.InsertCommand.Parameters.Add("p_unit_of_measure", OracleType.VarChar, 50)
                oda_l.InsertCommand.Parameters.Add("p_ship_to_organization_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_routing_header_id", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_aml_status", OracleType.Int32)
                oda_l.InsertCommand.Parameters.Add("p_comments", OracleType.VarChar, 240)
                oda_l.InsertCommand.Parameters.Add("p_delivery_type", OracleType.VarChar, 240)
                oda_l.InsertCommand.Parameters.Add("p_coo_code", OracleType.VarChar, 50)
                oda_l.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda_l.InsertCommand.Parameters.Add("o_return_message", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput

                oda_l.InsertCommand.Parameters("p_header_interface_id").SourceColumn = "header_interface_id"
                oda_l.InsertCommand.Parameters("p_user_id").SourceColumn = "user_id"
                oda_l.InsertCommand.Parameters("p_item_id").SourceColumn = "item_id"
                oda_l.InsertCommand.Parameters("p_item_category_id").SourceColumn = "item_category_id"
                oda_l.InsertCommand.Parameters("p_item_revision").SourceColumn = "item_revision"
                oda_l.InsertCommand.Parameters("p_quantity").SourceColumn = "REC_QTY"
                oda_l.InsertCommand.Parameters("p_unit_of_measure").SourceColumn = "po_uom"
                oda_l.InsertCommand.Parameters("p_ship_to_organization_id").SourceColumn = "ship_to_organization_id"
                oda_l.InsertCommand.Parameters("p_routing_header_id").SourceColumn = "receiving_routing_id"
                oda_l.InsertCommand.Parameters("p_aml_status").SourceColumn = "AML_STATUS"
                oda_l.InsertCommand.Parameters("p_comments").SourceColumn = "comments"
                oda_l.InsertCommand.Parameters("p_delivery_type").SourceColumn = "delivery_type"
                oda_l.InsertCommand.Parameters("p_coo_code").SourceColumn = "COO"
                oda_l.InsertCommand.Parameters("o_succ_flag").SourceColumn = "succ_flag"
                oda_l.InsertCommand.Parameters("o_return_message").SourceColumn = "return_message"

                oda_l.InsertCommand.Connection.Open()
                oda_l.Update(p_ds.Tables("po_line"))
                oda_l.InsertCommand.Connection.Close()

                Dim i As Integer
                Dim ErrMsg As String = ""
                Dim DRPO() As DataRow = Nothing
                Dim DRLine() As DataRow = Nothing

                DRPO = p_ds.Tables("po_header").Select("succ_flag = 'N' or succ_flag = ' '")
                DRLine = p_ds.Tables("po_line").Select("succ_flag = 'N' or succ_flag = ' '")
                If DRPO.Length > 0 OrElse DRLine.Length > 0 Then
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Receipt Error"

                    del_rcv_inte(LoginData.UserID, p_batch_id)
                    If DRLine.Length = 0 Then Return GRResult

                    'Record and Return Error message
                    For i = 0 To DRLine.Length - 1
                        If DRLine(i)("return_message").ToString.Trim <> "" Then
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = DRLine(i)("ITEM_NUMBER").ToString
                            ErrorRow(1) = DRLine(i)("return_message").ToString.Trim
                            ErrorTable.Rows.Add(ErrorRow)
                            ErrMsg = ErrMsg & "; " & DRLine(i)("ITEM_NUMBER").ToString & ": " & DRLine(i)("return_message").ToString
                        End If
                    Next
                    ErrorLogging("Receiving-Post_Receive1", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & Header.Type & " and return message; " & ErrMsg, "I")
                    GRResult.GRMsg = New DataSet
                    GRResult.GRMsg.Tables.Add(ErrorTable)
                    Return GRResult
                End If

                'Sometimes, we need to write PO data into interface table only for test  -- 08/10/2016
                'Submit Receipt Request to Oracle or not ? ( YES = Submit data to Oracle, NO =Not Submit) 
                Dim Sqlstr As String
                Dim SubmitFlag As String = ""
                Sqlstr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC009'")
                SubmitFlag = Convert.ToString(da.ExecuteScalar(Sqlstr)).ToUpper
                If SubmitFlag = "NO" Then
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Not Submit Receipt Request to Oracle -- Testing only"
                    Return GRResult
                End If


                Return Rcv_Submit(LoginData, p_batch_id, POStr, Header.Type)

            Catch oe As Exception
                ErrorLogging("Receiving-Post_Receive", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & Header.Type & ", " & oe.Message & oe.Source, "E")
                del_rcv_inte(LoginData.UserID, p_batch_id)
                Return Nothing
            Finally
                If oda_l.InsertCommand.Connection.State <> ConnectionState.Closed Then oda_l.InsertCommand.Connection.Close()
                If oda_h.InsertCommand.Connection.State <> ConnectionState.Closed Then oda_h.InsertCommand.Connection.Close()
            End Try
        End Using

    End Function

    Public Function Post_OSP_PO(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet, ByVal p_batch_id As Integer, ByVal Header As GRHeaderStructure) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Dim GRResult As CreateGRResponse = New CreateGRResponse

            Dim ErrorTable As DataTable
            Dim ErrorRow As Data.DataRow
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ColumnName", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim oda As OracleDataAdapter = da.Oda_Insert()          'da.Oda_Insert_Tran()
            Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables("po_line").Rows(0)("PO_NUMBER"), 4, 2)

            Try

                Dim Global_RespID As String = 0
                If POStr = "04" Then Global_RespID = GetGlobalRespID(LoginData.Server)

                oda.InsertCommand.CommandType = CommandType.StoredProcedure

                oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_header_osp"
                oda.InsertCommand.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
                oda.InsertCommand.Parameters.Add("p_header_interface_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_headers_comments", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_waybill_airbill_num", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_bill_of_lading", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("p_packing_slip", OracleType.VarChar, 240)
                oda.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters("p_header_interface_id").SourceColumn = "header_interface_id"
                oda.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                oda.InsertCommand.Parameters("p_user_id").SourceColumn = "user_id"
                oda.InsertCommand.Parameters("p_headers_comments").SourceColumn = "comments"
                oda.InsertCommand.Parameters("p_waybill_airbill_num").SourceColumn = "waybill_airbill_num"
                oda.InsertCommand.Parameters("p_bill_of_lading").SourceColumn = "bill_of_lading"
                oda.InsertCommand.Parameters("p_packing_slip").SourceColumn = "packing_slip"
                oda.InsertCommand.Parameters("o_succ_flag").SourceColumn = "succ_flag"

                oda.InsertCommand.Connection.Open()
                oda.Update(p_ds.Tables("po_header"))
                oda.InsertCommand.Parameters.Clear()

                oda.InsertCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.ins_rcv_tran_inte_osp_tran"
                oda.InsertCommand.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
                oda.InsertCommand.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
                oda.InsertCommand.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID
                oda.InsertCommand.Parameters.Add("p_global_resp_id", OracleType.Int32).Value = Global_RespID

                oda.InsertCommand.Parameters.Add("p_header_interface_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_po_header_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_po_release_id", OracleType.Int32)        'p_po_release_id
                oda.InsertCommand.Parameters.Add("p_vendor_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_vendor_site_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_user_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_po_line_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_po_line_location_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_po_distribution_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_item_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_item_category_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_item_revision", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_quantity", OracleType.[Double])
                oda.InsertCommand.Parameters.Add("p_unit_of_measure", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_line_num", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_ship_to_organization_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_routing_header_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_wip_entity_id", OracleType.Int32)
                oda.InsertCommand.Parameters.Add("p_assembly_reversion", OracleType.VarChar, 50)
                oda.InsertCommand.Parameters.Add("p_comments", OracleType.VarChar, 240)

                oda.InsertCommand.Parameters.Add("x_subinventory", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("x_locator", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_succ_flag", OracleType.VarChar, 50).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_return_message", OracleType.VarChar, 2000).Direction = ParameterDirection.InputOutput
                oda.InsertCommand.Parameters.Add("o_rcv_message", OracleType.VarChar, 8000).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_gr", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                oda.InsertCommand.Parameters.Add("o_rt_date", OracleType.VarChar, 50).Direction = ParameterDirection.Output

                oda.InsertCommand.Parameters("p_header_interface_id").SourceColumn = "header_interface_id"
                oda.InsertCommand.Parameters("p_po_header_id").SourceColumn = "po_header_id"
                oda.InsertCommand.Parameters("p_po_release_id").SourceColumn = "po_release_id"
                oda.InsertCommand.Parameters("p_vendor_id").SourceColumn = "vendor_id"
                oda.InsertCommand.Parameters("p_vendor_site_id").SourceColumn = "vendor_site_id"
                oda.InsertCommand.Parameters("p_user_id").SourceColumn = "user_id"
                oda.InsertCommand.Parameters("p_po_line_id").SourceColumn = "po_line_id"
                oda.InsertCommand.Parameters("p_po_line_location_id").SourceColumn = "po_line_location_id"
                oda.InsertCommand.Parameters("p_po_distribution_id").SourceColumn = "po_distribution_id"
                oda.InsertCommand.Parameters("p_item_id").SourceColumn = "item_id"
                oda.InsertCommand.Parameters("p_item_category_id").SourceColumn = "item_category_id"
                oda.InsertCommand.Parameters("p_item_revision").SourceColumn = "item_revision"
                oda.InsertCommand.Parameters("p_quantity").SourceColumn = "REC_QTY"
                oda.InsertCommand.Parameters("p_unit_of_measure").SourceColumn = "po_uom"
                oda.InsertCommand.Parameters("p_line_num").SourceColumn = "line_num"
                oda.InsertCommand.Parameters("p_ship_to_organization_id").SourceColumn = "ship_to_organization_id"
                oda.InsertCommand.Parameters("p_routing_header_id").SourceColumn = "receiving_routing_id"
                oda.InsertCommand.Parameters("p_wip_entity_id").SourceColumn = "wip_entity_id"
                oda.InsertCommand.Parameters("p_assembly_reversion").SourceColumn = "assembly_reversion"
                oda.InsertCommand.Parameters("p_comments").SourceColumn = "comments"

                oda.InsertCommand.Parameters("x_subinventory").SourceColumn = "WIP_COMPLETION_SUBINVENTORY"
                oda.InsertCommand.Parameters("x_locator").SourceColumn = "WIP_COMPLETION_LOCATOR"
                oda.InsertCommand.Parameters("o_succ_flag").SourceColumn = "succ_flag"
                oda.InsertCommand.Parameters("o_return_message").SourceColumn = "return_message"
                oda.InsertCommand.Parameters("o_rcv_message").SourceColumn = "rcv_message"
                oda.InsertCommand.Parameters("o_gr").SourceColumn = "gr"
                oda.InsertCommand.Parameters("o_rt_date").SourceColumn = "rt_date"

                oda.Update(p_ds.Tables("po_line"))
                oda.InsertCommand.Connection.Close()
                'oda.Dispose()

                Dim i, j As Integer
                Dim RT_Return As Boolean = False

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables("po_header").Select("succ_flag = 'Y' or succ_flag = 'S' ")
                If DR.Length > 0 Then RT_Return = True

                For i = 0 To p_ds.Tables("po_line").Rows.Count - 1
                    Dim Flag As String = p_ds.Tables("po_line").Rows(i)("succ_flag").ToString
                    If RT_Return = True Then
                        If Flag = "Y" Or Flag = "S" Then
                        Else
                            RT_Return = False
                        End If
                    End If

                    If RT_Return = False Then          '                        'Return Error message
                        GRResult.GRNo = ""
                        GRResult.GRStatus = "E"
                        GRResult.GRMessage = "Receipt Error"

                        Dim DJ As String = "DJ " & p_ds.Tables("po_line").Rows(i)("WIP_ENTITY_NAME").ToString
                        Dim ReturnMsg As String = p_ds.Tables("po_line").Rows(i)("return_message").ToString
                        If ReturnMsg.Trim <> "" Then
                            Dim MsgArry() As String
                            MsgArry = Split(ReturnMsg, ";")
                            For j = 0 To MsgArry.Length - 1
                                If MsgArry(j).ToString.Trim <> "" Then
                                    ErrorRow = ErrorTable.NewRow()
                                    ErrorRow(0) = p_ds.Tables("po_line").Rows(i)("WIP_ASSEMBLY") & " with return message "
                                    ErrorRow(1) = DJ & ", " & MsgArry(j).ToString.Trim
                                    ErrorTable.Rows.Add(ErrorRow)
                                End If
                            Next
                            Dim ErrMsg As String = "BatchID: " & p_batch_id & " for " & DJ & " with return message: " & ReturnMsg
                            ErrorLogging("Receiving-Post_OSP_PO1", LoginData.User.ToUpper, ErrMsg, "I")
                        End If
                    End If
                Next

                'Delete error from interface table after eTrace recorded it in errorlog table
                If RT_Return = False Then
                    del_rcv_inte(LoginData.UserID, p_batch_id)
                End If

                If ErrorTable.Rows.Count > 0 Then
                    GRResult.GRMsg = New DataSet
                    GRResult.GRMsg.Tables.Add(ErrorTable)
                    Return GRResult
                End If
                If GRResult.GRStatus = "E" Then Return GRResult

                'Return OSP_Submit(LoginData, p_batch_id, POStr)

                'Call OSP Submit the same as Normal PO Submit ---- Yudy 07/05/2012
                Return Rcv_Submit(LoginData, p_batch_id, POStr, Header.Type)

            Catch oe As Exception
                ErrorLogging("Receiving-Post_OSP_PO", LoginData.User.ToUpper, "BatchID: " & p_batch_id & ", " & oe.Message & oe.Source, "E")
                del_rcv_inte(LoginData.UserID, p_batch_id)
                Return Nothing
            Finally
                If oda.InsertCommand.Connection.State <> ConnectionState.Closed Then oda.InsertCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Rcv_Submit(ByVal LoginData As ERPLogin, ByVal p_batch_id As Integer, ByVal POStr As String, ByVal ReceiptType As Integer) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Dim GRResult As CreateGRResponse = New CreateGRResponse

            Dim ErrorTable As DataTable
            Dim ErrorRow As Data.DataRow
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ColumnName", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim Oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim Global_RespID As String = 0
                If POStr = "04" Then Global_RespID = GetGlobalRespID(LoginData.Server)

                Dim dsRTData As New DataSet()
                dsRTData.Tables.Add("msg")

                Dim comm_submit As OracleCommand = da.OraCommand()
                Oda.SelectCommand.CommandType = CommandType.StoredProcedure
                Oda.SelectCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.rcv_submit_req"
                'If ReceiptType = 2 Then
                '    Oda.SelectCommand.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.rcv_submit_osp_newreq"
                'End If

                Oda.SelectCommand.Parameters.Add("o_error_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters.Add("p_user_id", OracleType.Int32).Value = LoginData.UserID
                Oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OrgID  'LoginData.OrgCode
                Oda.SelectCommand.Parameters.Add("p_receipt_type", OracleType.Int32).Value = ReceiptType
                Oda.SelectCommand.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
                Oda.SelectCommand.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID
                Oda.SelectCommand.Parameters.Add("p_global_resp_id", OracleType.Int32).Value = Global_RespID
                Oda.SelectCommand.Parameters.Add("p_batch_id", OracleType.Int32).Value = p_batch_id
                Oda.SelectCommand.Parameters.Add("p_return_rcv_num", OracleType.VarChar).Value = "Y"
                Oda.SelectCommand.Parameters.Add("errbuff", OracleType.VarChar, 4000)
                Oda.SelectCommand.Parameters.Add("retcode", OracleType.VarChar, 240)

                Oda.SelectCommand.Parameters("errbuff").Direction = ParameterDirection.Output
                Oda.SelectCommand.Parameters("retcode").Direction = ParameterDirection.Output
                Oda.SelectCommand.Connection.Open()
                Oda.Fill(dsRTData, "msg")
                Oda.SelectCommand.Connection.Close()

                Dim i As Integer
                Dim ErrMsg As String = ""

                'Record and return Error message if the table is not we designed in oracle package
                If dsRTData.Tables(0).Columns(0).ColumnName.ToUpper <> "MSG_TYPE" Then
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Receipt Error"

                    If dsRTData.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dsRTData.Tables(0).Rows.Count - 1
                            ErrorRow = ErrorTable.NewRow()
                            If dsRTData.Tables(0).Columns.Count < 2 Then
                                ErrorRow(0) = "Error"
                                ErrorRow(1) = dsRTData.Tables(0).Rows(i)(0).ToString
                                ErrMsg = ErrMsg & "; " & dsRTData.Tables(0).Rows(i)(0).ToString
                            Else
                                ErrorRow(0) = dsRTData.Tables(0).Rows(i)(0).ToString
                                ErrorRow(1) = dsRTData.Tables(0).Rows(i)(1).ToString
                                ErrMsg = ErrMsg & "; " & dsRTData.Tables(0).Rows(i)(0).ToString & ": " & dsRTData.Tables(0).Rows(i)(1).ToString
                            End If
                            ErrorTable.Rows.Add(ErrorRow)
                        Next
                        ErrorLogging("Receiving-Rcv_Submit1", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & ReceiptType & " and error message; " & ErrMsg, "I")
                        GRResult.GRMsg = New DataSet
                        GRResult.GRMsg.Tables.Add(ErrorTable)
                    End If

                    'Delete error from interface table after eTrace recorded it in errorlog
                    del_rcv_inte(LoginData.UserID, p_batch_id)
                    Return GRResult
                End If

                Dim DR() As DataRow = Nothing
                DR = dsRTData.Tables(0).Select("msg_type = 'ERROR'")
                If DR.Length > 0 Then
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Receipt Error"

                    For i = 0 To DR.Length - 1
                        If DR(i)("text").ToString <> "" Then
                            ErrorRow = ErrorTable.NewRow()
                            ErrorRow(0) = DR(i)("msg_id").ToString
                            ErrorRow(1) = DR(i)("text").ToString
                            ErrorTable.Rows.Add(ErrorRow)
                            ErrMsg = ErrMsg & "; " & DR(i)("msg_id").ToString & ": " & DR(i)("text").ToString
                        End If
                    Next
                    ErrorLogging("Receiving-Rcv_Submit2", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & ReceiptType & " and error message; " & ErrMsg, "I")
                    GRResult.GRMsg = New DataSet
                    GRResult.GRMsg.Tables.Add(ErrorTable)

                    'Delete error from interface table after eTrace recorded it in errorlog
                    del_rcv_inte(LoginData.UserID, p_batch_id)
                    Return GRResult
                End If

                DR = Nothing
                DR = dsRTData.Tables(0).Select("msg_type = 'GR_Num'")
                If DR.Length > 0 Then
                    Dim GRNo As String = DR(0)("msg_id").ToString
                    If GRNo <> "" Then
                        GRResult.GRNo = GRNo
                        GRResult.PostDate = CDate(DR(0)("text"))                   'Get Receipt Date from Oracle after RT generated
                        GRResult.GRStatus = "Y"                                    'Receipt No created 
                        GRResult.GRMessage = "Receipt full completed"
                        GRResult.GRMsg = New DataSet
                        GRResult.GRMsg = dsRTData
                    Else
                        GRResult.GRNo = ""
                        GRResult.GRStatus = "E"                                    'Receipt Error
                        GRResult.GRMessage = DR(0)("text").ToString
                        ErrMsg = DR(0)("text").ToString

                        Dim ErrorRequest As String = "Program exited with status 1"
                        If ErrMsg.Contains(ErrorRequest) Then
                            GRResult.GRMessage = ErrorRequest & ". " & "Please contact your local IT."
                        End If
                        ErrorLogging("Receiving-Rcv_Submit3", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & ReceiptType & " and error message; " & ErrMsg, "I")
                    End If
                Else
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Receipt Error"
                End If

                'Delete error from interface table after eTrace recorded it in errorlog
                If GRResult.GRStatus = "E" Then
                    del_rcv_inte(LoginData.UserID, p_batch_id)
                End If

                Return GRResult

            Catch oe As Exception
                ErrorLogging("Receiving-Rcv_Submit", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with Receipt type " & ReceiptType & ",  " & oe.Message & oe.Source, "E")
                del_rcv_inte(LoginData.UserID, p_batch_id)
                Return Nothing
            Finally
                If Oda.SelectCommand.Connection.State <> ConnectionState.Closed Then Oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function OSP_Submit(ByVal LoginData As ERPLogin, ByVal p_batch_id As Integer, ByVal POStr As String) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Dim GRResult As CreateGRResponse = New CreateGRResponse

            Dim ErrorTable As DataTable
            Dim ErrorRow As Data.DataRow
            Dim myDataColumn As DataColumn
            ErrorTable = New Data.DataTable("ErrorTable")
            myDataColumn = New Data.DataColumn("ColumnName", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
            ErrorTable.Columns.Add(myDataColumn)

            Dim aa As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                Dim OrgID As String = GetOrgID(LoginData.OrgCode)

                Dim Global_RespID As String = 0
                If POStr = "04" Then Global_RespID = GetGlobalRespID(LoginData.Server)

                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.rcv_submit_osp"
                OC.Parameters.Add("p_user_id", OracleType.Int32).Value = LoginData.UserID
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OrgID  'LoginData.OrgCode
                OC.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv
                OC.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv
                OC.Parameters.Add("p_global_resp_id", OracleType.Int32).Value = Global_RespID
                OC.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id

                OC.Parameters.Add("o_succ_flag", OracleType.VarChar, 100).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_return_message", OracleType.VarChar, 5000).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_rcv_message", OracleType.VarChar, 5000).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_gr", OracleType.VarChar, 100).Direction = ParameterDirection.Output
                OC.Parameters.Add("o_rt_date", OracleType.VarChar, 100).Direction = ParameterDirection.Output

                If OC.Connection.State = ConnectionState.Closed Then OC.Connection.Open()
                aa = CInt(OC.ExecuteNonQuery())

                Dim Flag, RcvMsg, ReturnMsg, RTNo, RTDate As String
                Flag = OC.Parameters("o_succ_flag").Value.ToString
                RTNo = OC.Parameters("o_gr").Value.ToString
                RTDate = OC.Parameters("o_rt_date").Value.ToString
                RcvMsg = OC.Parameters("o_return_message").Value.ToString
                ReturnMsg = OC.Parameters("o_rcv_message").Value.ToString

                OC.Connection.Close()
                'OC.Dispose()

                If Flag = "Y" OrElse Flag = "S" Then
                    If RTNo.ToString <> "" Then
                        GRResult.GRNo = RTNo.ToString
                        GRResult.GRStatus = "Y"                             'Receipt No created 
                        GRResult.GRMessage = "Receipt full completed"
                        GRResult.PostDate = CDate(RTDate)                   'Get Receipt Date from Oracle after RT generated
                    Else
                        GRResult.GRNo = ""
                        GRResult.GRStatus = "E"
                        GRResult.GRMessage = RcvMsg.ToString.Trim
                    End If

                Else
                    GRResult.GRNo = ""
                    GRResult.GRStatus = "E"
                    GRResult.GRMessage = "Receipt Error"

                    If RcvMsg.ToString <> "" Then
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = "RCV Message "
                        ErrorRow(1) = RcvMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        ErrorLogging("Receiving-OSP_Submit-RcvMsg", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with RCV message: " & RcvMsg, "I")
                    End If
                    If ReturnMsg.ToString <> "" Then
                        ErrorRow = ErrorTable.NewRow()
                        ErrorRow(0) = "Return Message "
                        ErrorRow(1) = ReturnMsg
                        ErrorTable.Rows.Add(ErrorRow)
                        ErrorLogging("Receiving-OSP_Submit-ReturnMsg", LoginData.User.ToUpper, "BatchID: " & p_batch_id & " with return message: " & ReturnMsg, "I")
                    End If
                    GRResult.GRMsg = New DataSet
                    GRResult.GRMsg.Tables.Add(ErrorTable)

                End If

                'Delete error from interface table after eTrace recorded it in errorlog table
                If GRResult.GRStatus = "E" Then
                    del_rcv_inte(LoginData.UserID, p_batch_id)
                End If

                Return GRResult

            Catch oe As Exception
                ErrorLogging("Receiving-OSP_Submit", LoginData.User.ToUpper, "BatchID: " & p_batch_id & ", " & oe.Message & oe.Source, "E")
                del_rcv_inte(LoginData.UserID, p_batch_id)
                Return Nothing
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
        End Using
    End Function

    Public Function del_rcv_inte(ByVal p_user_id As Int32, ByVal p_batch_id As Integer) As String
        'del_rcv_inte(p_group_id number,p_user_id number,o_succ_flag out varchar2,o_return_message out varchar2)
        Using da As DataAccess = GetDataAccess()

            Dim aa As Integer
            Dim MsgFlag As String
            Dim Oda As OracleCommand = da.OraCommand()
            Try
                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.xxetr_inv_rcv_inte_hand_pkg.del_rcv_inte"
                Oda.Parameters.Add("p_group_id", OracleType.Int32).Value = p_batch_id
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = p_user_id
                'Oda.Parameters.Add("p_org_code", OracleType.VarChar, 240).Value = OrgCode
                Oda.Parameters.Add("o_succ_flag", OracleType.VarChar, 50)
                Oda.Parameters.Add("o_return_message", OracleType.VarChar, 240)
                Oda.Parameters("o_succ_flag").Direction = ParameterDirection.Output
                Oda.Parameters("o_return_message").Direction = ParameterDirection.Output

                Oda.Connection.Open()
                aa = CInt(Oda.ExecuteNonQuery())
                MsgFlag = Oda.Parameters("o_succ_flag").Value
                Oda.Connection.Close()
                Return DirectCast(MsgFlag, String)

            Catch oe As Exception
                ErrorLogging("Receiving-del_rcv_inte", "", "BatchID: " & p_batch_id & ", " & oe.Message & oe.Source, "E")
                Return "N"
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    'Private Function CancelGR(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet) As CreateGRResponse
    '    Using da As DataAccess = GetDataAccess()

    '        Try

    '            Dim ds As DataSet = New DataSet
    '            Dim RTData As DataTable
    '            Dim myDataColumn As DataColumn
    '            Dim myDataRow As Data.DataRow

    '            RTData = New Data.DataTable("RTData")
    '            myDataColumn = New Data.DataColumn("RTNo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("POLineID", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("PONo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("ReleaseNo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("POLine", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("ShipmentNo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.Decimal"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("BaseUOM", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("SubInventory", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("ReasonCode", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("RMANo", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("KanbanCard", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("POType", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)
    '            myDataColumn = New Data.DataColumn("Message", System.Type.GetType("System.String"))
    '            RTData.Columns.Add(myDataColumn)

    '            ds.Tables.Add(RTData)

    '            Dim dtDashB As DataTable = New DataTable
    '            dtDashB.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
    '            dtDashB.Columns.Add(New Data.DataColumn("Packages", System.Type.GetType("System.String")))
    '            dtDashB.Columns.Add(New Data.DataColumn("RecQty", System.Type.GetType("System.Decimal")))

    '            Dim i, j As Integer
    '            Dim DR(), PORow(), CLIDRow() As DataRow
    '            Dim PONo, POItem, ItemText, CLID, CurrID As String
    '            Dim POLineID, MaterialNo, SubInventory, Locator As String

    '            Dim SqlStr As String
    '            Dim MatlLists As String = ""
    '            Dim RecConfig As String = ""

    '            Try
    '                'Read Receiving DashBoard Flag from Table T_Config to decide if need to Delete PalletID info from table T_RecDashboard 
    '                SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC001'")
    '                RecConfig = Convert.ToString(da.ExecuteScalar(SqlStr))

    '                Dim CLIDTable As DataTable = New DataTable
    '                CLIDTable = Items.Tables(1).Copy

    '                CLIDRow = Nothing
    '                CLIDRow = CLIDTable.Select(" Status = 'X' and RefCLID IS NOT NULL and RefCLID <> ' ' ")
    '                If CLIDRow.Length > 0 Then
    '                    For i = 0 To CLIDRow.Length - 1
    '                        Dim RefCLID As String = CLIDRow(i)("RefCLID").ToString
    '                        DR = Nothing
    '                        DR = Items.Tables(1).Select(" RefCLID = '" & RefCLID & "'")
    '                        If DR.Length > 0 Then
    '                            For j = 0 To DR.Length - 1
    '                                If DR(j)("Status").ToString = "" Then
    '                                    DR(j)("Status") = "X"
    '                                    DR(j).AcceptChanges()
    '                                End If
    '                            Next
    '                            Items.Tables(1).AcceptChanges()
    '                        End If
    '                    Next
    '                End If


    '                For i = 0 To Items.Tables(1).Rows.Count - 1
    '                    If Items.Tables(1).Rows(i)("Status") Is DBNull.Value OrElse Items.Tables(1).Rows(i)("Status") = "" Then
    '                    ElseIf Items.Tables(1).Rows(i)("Status") = "X" Then

    '                        PONo = Items.Tables(1).Rows(i)("PONO").ToString
    '                        POItem = Items.Tables(1).Rows(i)("POItem").ToString
    '                        POLineID = Items.Tables(1).Rows(i)("POLineID").ToString
    '                        SubInventory = Items.Tables(1).Rows(i)("SubInventory").ToString
    '                        Locator = Items.Tables(1).Rows(i)("Locator").ToString

    '                        PORow = Nothing
    '                        PORow = Items.Tables(0).Select(" PONO = '" & PONo & "' and POItem = '" & POItem & "' and POLineID = '" & POLineID & "'")
    '                        If PORow.Length > 0 Then
    '                            If PORow(0)("PullMethod").ToString <> "" Then PORow(0)("ReasonCode") = ""
    '                            'POLineID = PORow(0)("POLineID").ToString

    '                            DR = Nothing
    '                            DR = ds.Tables("RTData").Select(" POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                            If DR.Length = 0 Then
    '                                myDataRow = ds.Tables("RTData").NewRow()
    '                                myDataRow("RTNo") = Header.OrderNo
    '                                myDataRow("MaterialNo") = Items.Tables(1).Rows(i)("MatNo")
    '                                myDataRow("POLineID") = POLineID

    '                                Dim OrderData As POData
    '                                OrderData = Split_POData(PONo, POItem)

    '                                myDataRow("PONo") = OrderData.PONo
    '                                myDataRow("POLine") = OrderData.LineNo
    '                                myDataRow("ShipmentNo") = OrderData.ShipmentNo
    '                                If OrderData.ReleaseNo = 0 Then
    '                                    myDataRow("ReleaseNo") = ""
    '                                Else
    '                                    myDataRow("ReleaseNo") = OrderData.ReleaseNo
    '                                End If

    '                                myDataRow("QtyBaseUOM") = Convert.ToDecimal(Items.Tables(1).Rows(i)("Qty"))
    '                                myDataRow("BaseUOM") = PORow(0)("BaseUOM").ToString
    '                                myDataRow("SubInventory") = SubInventory
    '                                myDataRow("Locator") = Locator
    '                                myDataRow("ReasonCode") = PORow(0)("ReasonCode").ToString
    '                                myDataRow("RMANo") = PORow(0)("ItemText").ToString

    '                                If PORow(0)("PullMethod").ToString = "K" Then
    '                                    myDataRow("KanbanCard") = POLineID
    '                                End If
    '                                myDataRow("POType") = PORow(0)("PullMethod").ToString
    '                                myDataRow("Status") = ""
    '                                myDataRow("Message") = ""
    '                                ds.Tables("RTData").Rows.Add(myDataRow)
    '                            Else
    '                                Dim QtyBaseUOM As Decimal
    '                                QtyBaseUOM = DR(0)("QtyBaseUOM")

    '                                DR(0)("QtyBaseUOM") = QtyBaseUOM + Items.Tables(1).Rows(i)("Qty")
    '                                DR(0).AcceptChanges()
    '                                DR(0).SetAdded()
    '                            End If

    '                            Items.Tables(0).AcceptChanges()
    '                        End If
    '                    End If
    '                Next


    '                Dim ReturnData As DataSet = New DataSet

    '                Try
    '                    If ds.Tables(0).Rows.Count > 0 Then
    '                        ReturnData = Oracle_Return(LoginData, ds)
    '                    End If
    '                Catch ex As Exception
    '                    ErrorLogging("Receiving-CancelGR1", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                    ReturnData = Nothing
    '                End Try

    '                If ReturnData Is Nothing OrElse ReturnData.Tables.Count = 0 OrElse ReturnData.Tables(0).Rows.Count = 0 Then
    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                    Exit Function
    '                End If


    '                Dim ErrorTable As DataTable
    '                Dim ErrorRow As Data.DataRow
    '                ErrorTable = New Data.DataTable("ErrorTable")
    '                myDataColumn = New Data.DataColumn("ColumnName", System.Type.GetType("System.String"))
    '                ErrorTable.Columns.Add(myDataColumn)
    '                myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
    '                ErrorTable.Columns.Add(myDataColumn)


    '                DR = Nothing
    '                DR = ReturnData.Tables(0).Select("Status = 'N' or Status = ' '")
    '                If DR.Length > 0 Then
    '                    For i = 0 To DR.Length - 1
    '                        If DR(i)("Message").ToString <> "" Then
    '                            ErrorRow = ErrorTable.NewRow()
    '                            ErrorRow("ColumnName") = DR(i)("MaterialNo")
    '                            ErrorRow("ErrorMsg") = DR(i)("Message").ToString
    '                            ErrorTable.Rows.Add(ErrorRow)
    '                        End If
    '                    Next
    '                    CancelGR.GRMsg = New DataSet
    '                    CancelGR.GRMsg.Tables.Add(ErrorTable)

    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                End If


    '                DR = Nothing
    '                DR = ReturnData.Tables(0).Select("Status = 'Y' ")
    '                If DR.Length = 0 Then
    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                    Exit Function
    '                End If


    '                For i = 0 To DR.Length - 1
    '                    MaterialNo = DR(i)("MaterialNo").ToString
    '                    SubInventory = DR(i)("SubInventory").ToString
    '                    Locator = DR(i)("Locator").ToString
    '                    POLineID = DR(i)("POLineID").ToString

    '                    PORow = Nothing
    '                    PORow = Items.Tables(0).Select(" POLineID = '" & POLineID & "'")
    '                    If PORow.Length = 0 Then Continue For

    '                    CurrID = ""
    '                    PONo = PORow(0)("PONO")
    '                    POItem = PORow(0)("POItem")
    '                    ItemText = ReturnData.Tables(0).Rows(i)("RMANo")

    '                    CLIDRow = Nothing
    '                    'CLIDRow = Items.Tables(1).Select("Status = 'X' and PONO = '" & PONo & "' and POItem = '" & POItem & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                    CLIDRow = Items.Tables(1).Select("Status = 'X' and POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                    If CLIDRow.Length = 0 Then Continue For

    '                    Dim NoOfPackage As Integer = CLIDRow.Length
    '                    For j = 0 To CLIDRow.Length - 1
    '                        If CLIDRow(j)("RefCLID") Is DBNull.Value Then CLIDRow(j)("RefCLID") = ""
    '                        If CLIDRow(j)("RefCLID") <> "" Then
    '                            CLID = CLIDRow(j)("RefCLID")
    '                        Else
    '                            CLID = CLIDRow(j)("LID")
    '                        End If

    '                        If CurrID = CLID Then
    '                            NoOfPackage = NoOfPackage - 1
    '                            Continue For
    '                        End If
    '                        CurrID = CLID

    '                        Dim StartStr As String = Microsoft.VisualBasic.Left(CLID, 1)
    '                        Dim MidStr As String = Microsoft.VisualBasic.Mid(CLID, 3, 1)

    '                        Try
    '                            'Update CLMaster table
    '                            If StartStr = "B" OrElse StartStr = "P" OrElse MidStr = "V" Then           'Delete record from T_CLMaster if the LabelID is CartonID / PalletID, or Vendor CLID
    '                                SqlStr = String.Format("DELETE from T_CLMaster Where CLID='{0}' ", CLID)
    '                            Else
    '                                SqlStr = String.Format("Update T_CLMaster set StatusCode=0, ChangedOn=getdate(), LastTransaction='Receipt Reversal', ChangedBy='{0}', AddlText='{1}', BoxID='{2}' Where CLID='{3}' ", LoginData.User.ToUpper, ItemText, DBNull.Value, CLID)
    '                            End If
    '                            da.ExecuteNonQuery(SqlStr)
    '                        Catch ex As Exception
    '                            ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                        End Try


    '                        Try
    '                            'Update eJIT table
    '                            Dim eJITID As String
    '                            Dim JITQty, tmpQty As Decimal

    '                            SqlStr = String.Format("Select SourceID from T_CLMaster with (nolock) WHERE ProcessID='302' and CLID ='{0}'", CLID)
    '                            eJITID = Convert.ToString(da.ExecuteScalar(SqlStr))

    '                            If eJITID <> "" Then
    '                                JITQty = CLIDRow(j)("Qty")

    '                                SqlStr = String.Format("Select RecQty FROM T_KBPickingList with (nolock) WHERE EJITID='{0}'", eJITID)
    '                                tmpQty = Convert.ToDecimal(da.ExecuteScalar(SqlStr))
    '                                If tmpQty > 0 Then
    '                                    tmpQty = tmpQty - JITQty
    '                                    SqlStr = String.Format("UPDATE T_KBPickingList set RecQty='{1}' where EJITID='{0}'", eJITID, tmpQty)
    '                                    da.ExecuteNonQuery(SqlStr)
    '                                End If
    '                            End If
    '                        Catch ex As Exception
    '                            ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                        End Try
    '                    Next

    '                    Dim POType As String = DR(i)("POType").ToString

    '                    'Record the Pallet info which needs to do Reverse as well
    '                    'POType: Blank(Normal PO), OSP(OSP PO), K(KB PO), EJIT, IR
    '                    If POType = "" AndAlso RecConfig = "YES" Then
    '                        Dim DBdr() As DataRow = Nothing
    '                        DBdr = dtDashB.Select(" MaterialNo = '" & MaterialNo & "'")
    '                        If DBdr.Length = 0 Then
    '                            Dim myDR As DataRow
    '                            myDR = dtDashB.NewRow()
    '                            myDR("MaterialNo") = MaterialNo
    '                            myDR("Packages") = NoOfPackage
    '                            myDR("RecQty") = DR(i)("QtyBaseUOM")
    '                            dtDashB.Rows.Add(myDR)

    '                            'Record the MaterialNo Lists which have done reversal
    '                            If MatlLists = "" Then
    '                                MatlLists = MaterialNo
    '                            Else
    '                                MatlLists = MatlLists & "," & MaterialNo
    '                            End If
    '                        Else
    '                            DBdr(0)("Packages") = DBdr(0)("Packages") + NoOfPackage
    '                            DBdr(0)("RecQty") = DBdr(0)("RecQty") + DR(i)("QtyBaseUOM")
    '                        End If
    '                    End If
    '                Next
    '            Catch ex As Exception
    '                ErrorLogging("Receiving-CancelGR3", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '            End Try


    '            'Save the Pallet info in table T_RecDashboard which needs to do Component Delivery
    '            If RecConfig = "YES" AndAlso dtDashB.Rows.Count > 0 Then

    '                MatlLists = "('" & MatlLists.Replace(",", "','") & "')"

    '                ds = New DataSet
    '                SqlStr = String.Format("Select MaterialNo,Packages,RecQty from T_RecDashboard with (nolock) WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo in {2}", LoginData.OrgCode, Header.OrderNo, MatlLists)
    '                ds = da.ExecuteDataSet(SqlStr, "DBItems")

    '                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
    '                ElseIf ds.Tables(0).Rows.Count > 0 Then
    '                    For i = 0 To dtDashB.Rows.Count - 1
    '                        MaterialNo = dtDashB.Rows(i)("MaterialNo").ToString

    '                        DR = Nothing
    '                        DR = ds.Tables(0).Select(" MaterialNo = '" & MaterialNo & "'")
    '                        If DR.Length = 0 Then Continue For

    '                        Dim RestQty As Decimal = DR(0)("RecQty") - dtDashB.Rows(i)("RecQty")

    '                        If RestQty > 0 Then
    '                            Dim Remarks As String = DateTime.Now.ToString & " by User: " & LoginData.User.ToUpper & " with LastTransaction: Update RecQty by Receipt Reversal"
    '                            SqlStr = String.Format("Update T_RecDashboard set RecQty='{0}',Remarks='{1}' WHERE OrgCode='{2}' and RTNo ='{3}' and MaterialNo ='{4}'", RestQty, Remarks, LoginData.OrgCode, Header.OrderNo, MaterialNo)
    '                        Else
    '                            SqlStr = String.Format("Delete from T_RecDashboard WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo ='{2}'", LoginData.OrgCode, Header.OrderNo, MaterialNo)
    '                        End If
    '                        da.ExecuteNonQuery(SqlStr)
    '                    Next
    '                End If
    '            End If


    '            ' Set Flag to show Partial complete or full complete
    '            If CancelGR.GRStatus Is Nothing Then
    '                CancelGR.GRStatus = "Y"
    '                CancelGR.GRMessage = "Reversal full completed "
    '            Else
    '                CancelGR.GRStatus = "N"
    '                CancelGR.GRMessage = "Reversal partial completed "
    '            End If
    '            CancelGR.GRNo = Header.OrderNo

    '            If Header.Type = 8 Then
    '                CancelGR.CLIDs = New DataSet
    '                CancelGR.CLIDs = Items
    '            End If

    '        Catch ex As Exception
    '            ErrorLogging("Receiving-CancelGR", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '        End Try

    '    End Using

    'End Function

    'Public Function Oracle_Return(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As DataSet
    '    Using da As DataAccess = GetDataAccess()

    '        Dim aa As OracleString
    '        Dim oc As OracleCommand = da.Ora_Command_Trans()
    '        Dim oda As OracleDataAdapter = New OracleDataAdapter()

    '        Dim RTNo As String = p_ds.Tables("RTData").Rows(0)("RTNo").ToString
    '        Dim POType As String = p_ds.Tables("RTData").Rows(0)("POType").ToString

    '        Try

    '            Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables("RTData").Rows(0)("PONo"), 4, 2)
    '            If POStr = "04" Then LoginData.RespID_Inv = GetGlobalRespID(LoginData.Server)


    '            oc.CommandType = CommandType.StoredProcedure
    '            oc.CommandText = "apps.xxetr_wip_pkg.initialize"
    '            oc.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
    '            oc.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
    '            oc.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID

    '            oc.ExecuteOracleNonQuery(aa)
    '            oc.Parameters.Clear()

    '            If POType = "OSP" Then
    '                oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return_osp"       'OSP PO Reversal
    '            Else
    '                oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return"               'Normal PO / KanBan PO Reversal
    '                oc.Parameters.Add("p_kanban_card_id", OracleType.VarChar, 500)
    '                oc.Parameters("p_kanban_card_id").SourceColumn = "KanbanCard"
    '            End If

    '            oc.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = LoginData.OrgCode
    '            oc.Parameters.Add("p_po_number", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_receipt_num", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
    '            oc.Parameters.Add("p_line_num", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_shipment_num", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_release_num", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_quantity", OracleType.Double)
    '            oc.Parameters.Add("p_uom_code", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
    '            oc.Parameters.Add("p_reason_code", OracleType.VarChar, 200).Value = ""    'Use default ReasonCode from Oracle
    '            oc.Parameters.Add("p_rma_number", OracleType.VarChar, 500)

    '            oc.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
    '            oc.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

    '            oc.Parameters("p_po_number").SourceColumn = "PONo"
    '            oc.Parameters("p_receipt_num").SourceColumn = "RTNo"
    '            oc.Parameters("p_transaction_id").SourceColumn = "POLineID"                      'TransactionID           -- 07/12/2017
    '            oc.Parameters("p_line_num").SourceColumn = "POLine"
    '            oc.Parameters("p_shipment_num").SourceColumn = "ShipmentNo"
    '            oc.Parameters("p_release_num").SourceColumn = "ReleaseNo"
    '            oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
    '            oc.Parameters("p_uom_code").SourceColumn = "BaseUOM"
    '            oc.Parameters("p_subinv").SourceColumn = "SubInventory"
    '            oc.Parameters("p_locator").SourceColumn = "Locator"
    '            'oc.Parameters("p_reason_code").SourceColumn = "ReasonCode"
    '            oc.Parameters("p_rma_number").SourceColumn = "RMANo"
    '            oc.Parameters("o_success_flag").SourceColumn = "Status"
    '            oc.Parameters("o_error_mssg").SourceColumn = "Message"

    '            oda.InsertCommand = oc
    '            oda.Update(p_ds.Tables("RTData"))

    '            Dim DR() As DataRow = Nothing
    '            DR = p_ds.Tables("RTData").Select("Status = 'N' or Status = ' '")
    '            If DR.Length = 0 Then
    '                oc.Transaction.Commit()
    '                oc.Connection.Close()
    '                'oc.Connection.Dispose()
    '                'oc.Dispose()
    '                Return p_ds
    '                Exit Function
    '            End If


    '            'Record error message to eTrace database
    '            Dim i As Integer
    '            Dim ErrMsg As String = ""
    '            For i = 0 To DR.Length - 1
    '                If DR(i)("Message").ToString.Trim <> "" Then
    '                    If ErrMsg = "" Then
    '                        ErrMsg = DR(i)("Message").ToString.Trim
    '                    Else
    '                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString.Trim
    '                    End If
    '                End If
    '            Next
    '            ErrorLogging("Receiving-Oracle_Return1", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with return message: " & ErrMsg, "I")

    '            oc.Transaction.Rollback()
    '            oc.Connection.Close()
    '            'oc.Connection.Dispose()
    '            'oc.Dispose()

    '            Return p_ds

    '        Catch oe As Exception
    '            ErrorLogging("Receiving-Oracle_Return", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with error message: " & oe.Message & oe.Source, "E")
    '            p_ds.Tables("RTData").Rows(0)("Status") = "N"
    '            p_ds.Tables("RTData").Rows(0)("Message") = oe.Message
    '            p_ds.Tables("RTData").Rows(0).AcceptChanges()

    '            If oc.Connection.State = ConnectionState.Open Then
    '                oc.Transaction.Rollback()
    '                oc.Connection.Close()
    '            End If
    '            Return p_ds
    '        Finally
    '            If oc.Connection.State <> ConnectionState.Closed Then oc.Connection.Close()
    '        End Try

    '    End Using
    'End Function

    Private Function CancelGR(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet) As CreateGRResponse
        Using da As DataAccess = GetDataAccess()

            Try

                Dim ds As DataSet = New DataSet
                Dim RTData As DataTable
                Dim myDataRow As Data.DataRow

                RTData = New Data.DataTable("RTData")
                RTData.Columns.Add(New Data.DataColumn("RTNo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("POLineID", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("PONo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("ReleaseNo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("POLine", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("ShipmentNo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("BaseUOM", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("SubInventory", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("Locator", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("RTLot", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("RMANo", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("POType", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("AddFlag", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("CorrectType", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
                RTData.Columns.Add(New Data.DataColumn("TranID", System.Type.GetType("System.String")))
                ds.Tables.Add(RTData)

                Dim dtDashB As DataTable = New DataTable
                dtDashB.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
                dtDashB.Columns.Add(New Data.DataColumn("Packages", System.Type.GetType("System.String")))
                dtDashB.Columns.Add(New Data.DataColumn("RecQty", System.Type.GetType("System.Decimal")))

                Dim dtDelivery As DataTable = New DataTable
                dtDelivery.Columns.Add(New Data.DataColumn("POLineID", System.Type.GetType("System.String")))
                dtDelivery.Columns.Add(New Data.DataColumn("Qty", System.Type.GetType("System.Decimal")))

                Dim i, j As Integer
                Dim DR(), PORow(), CLIDRow() As DataRow
                Dim PONo, POItem, ItemText, CLID, CurrID As String
                Dim POLineID, MaterialNo, SubInventory, Locator As String

                Dim SqlStr As String
                Dim MatlLists As String = ""
                Dim RecConfig As String = ""

                Try
                    'Read Receiving DashBoard Flag from Table T_Config to decide if need to Delete PalletID info from table T_RecDashboard 
                    SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC001'")
                    RecConfig = Convert.ToString(da.ExecuteScalar(SqlStr))

                    Dim CLIDTable As DataTable = New DataTable
                    CLIDTable = Items.Tables(1).Copy

                    CLIDRow = Nothing
                    CLIDRow = CLIDTable.Select(" Status = 'X' and RefCLID IS NOT NULL and RefCLID <> ' ' ")
                    If CLIDRow.Length > 0 Then
                        For i = 0 To CLIDRow.Length - 1
                            Dim RefCLID As String = CLIDRow(i)("RefCLID").ToString
                            DR = Nothing
                            DR = Items.Tables(1).Select(" RefCLID = '" & RefCLID & "'")
                            If DR.Length > 0 Then
                                For j = 0 To DR.Length - 1
                                    If DR(j)("Status").ToString = "" Then
                                        DR(j)("Status") = "X"
                                        DR(j).AcceptChanges()
                                    End If
                                Next
                                Items.Tables(1).AcceptChanges()
                            End If
                        Next
                    End If

                    Dim POType As String = Items.Tables(0).Rows(0)("PullMethod").ToString

                    For i = 0 To Items.Tables(1).Rows.Count - 1
                        If Items.Tables(1).Rows(i)("Status") Is DBNull.Value OrElse Items.Tables(1).Rows(i)("Status") = "" Then
                        ElseIf Items.Tables(1).Rows(i)("Status") = "X" Then

                            PONo = Items.Tables(1).Rows(i)("PONO").ToString
                            POItem = Items.Tables(1).Rows(i)("POItem").ToString
                            POLineID = Items.Tables(1).Rows(i)("POLineID").ToString
                            SubInventory = Items.Tables(1).Rows(i)("SubInventory").ToString
                            Locator = Items.Tables(1).Rows(i)("Locator").ToString

                            'For ASN Receipt, Collect Qty by TransactionID here for later Correction
                            'If POType <> "OSP" Then
                            If POType = "ASN" Then
                                PORow = Nothing
                                PORow = dtDelivery.Select(" POLineID = '" & POLineID & "'")
                                If PORow.Length = 0 Then
                                    myDataRow = dtDelivery.NewRow()
                                    myDataRow("POLineID") = POLineID
                                    myDataRow("Qty") = Convert.ToDecimal(Items.Tables(1).Rows(i)("Qty"))
                                    dtDelivery.Rows.Add(myDataRow)
                                Else
                                    Dim Qty As Decimal = PORow(0)("Qty")
                                    PORow(0)("Qty") = Qty + Items.Tables(1).Rows(i)("Qty")
                                    PORow(0).AcceptChanges()
                                End If

                            End If

                            PORow = Nothing
                            PORow = Items.Tables(0).Select(" PONO = '" & PONo & "' and POItem = '" & POItem & "' and POLineID = '" & POLineID & "'")
                            If PORow.Length > 0 Then
                                DR = Nothing
                                DR = ds.Tables("RTData").Select(" POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
                                If DR.Length = 0 Then
                                    myDataRow = ds.Tables("RTData").NewRow()
                                    myDataRow("RTNo") = Header.OrderNo
                                    myDataRow("MaterialNo") = Items.Tables(1).Rows(i)("MatNo")
                                    myDataRow("POLineID") = POLineID
                                    If Items.Tables(1).Rows(i)("SubInventory") = "" Then
                                        myDataRow("TranID") = Convert.ToInt32(da.ExecuteScalar(String.Format("exec dbo.ora_IQC_TransactionID {0}", POLineID)))
                                    Else
                                        myDataRow("TranID") = POLineID
                                    End If

                                    Dim OrderData As POData
                                    OrderData = Split_POData(PONo, POItem)

                                    myDataRow("PONo") = OrderData.PONo
                                    myDataRow("POLine") = OrderData.LineNo
                                    myDataRow("ShipmentNo") = OrderData.ShipmentNo
                                    If OrderData.ReleaseNo = 0 Then
                                        myDataRow("ReleaseNo") = ""
                                    Else
                                        myDataRow("ReleaseNo") = OrderData.ReleaseNo
                                    End If

                                    myDataRow("QtyBaseUOM") = Convert.ToDecimal(Items.Tables(1).Rows(i)("Qty"))
                                    myDataRow("BaseUOM") = PORow(0)("BaseUOM").ToString
                                    myDataRow("SubInventory") = SubInventory
                                    myDataRow("Locator") = Locator
                                    myDataRow("RTLot") = PORow(0)("RTLot").ToString
                                    myDataRow("RMANo") = PORow(0)("ItemText").ToString
                                    myDataRow("POType") = PORow(0)("PullMethod").ToString
                                    myDataRow("AddFlag") = ""
                                    myDataRow("Status") = ""
                                    myDataRow("Message") = ""
                                    ds.Tables("RTData").Rows.Add(myDataRow)
                                Else
                                    Dim QtyBaseUOM As Decimal
                                    QtyBaseUOM = DR(0)("QtyBaseUOM")

                                    DR(0)("QtyBaseUOM") = QtyBaseUOM + Items.Tables(1).Rows(i)("Qty")
                                    DR(0).AcceptChanges()
                                    DR(0).SetAdded()
                                End If

                                Items.Tables(0).AcceptChanges()
                            End If
                        End If
                    Next

                    'For ASN Receipt, Collect Qty by TransactionID here for later Correction
                    'If POType <> "OSP" Then
                    If POType = "ASN" Then
                        Dim dtRTData As New DataTable
                        dtRTData = ds.Tables("RTData").Clone

                        For i = 0 To dtDelivery.Rows.Count - 1
                            Dim TotalFlag As Boolean = False
                            Dim TotalQty As Decimal = dtDelivery.Rows(i)("Qty")

                            DR = Nothing
                            DR = ds.Tables("RTData").Select(" POLineID = '" & dtDelivery.Rows(i)("POLineID") & "'")
                            If DR.Length > 0 Then
                                For j = 0 To DR.Length - 1
                                    Dim mySubInv As String = DR(j)("SubInventory").ToString
                                    If mySubInv = "" Then
                                        TotalFlag = True
                                        DR(j)("CorrectType") = "TOTAL"
                                        DR(j)("QtyBaseUOM") = TotalQty
                                        dtRTData.ImportRow(DR(j))
                                    Else
                                        DR(j)("CorrectType") = "DELIVER"
                                        dtRTData.ImportRow(DR(j))
                                    End If
                                Next

                                'If All Qty has done Delivery, need to add one row which has Total Qty
                                If TotalFlag = False Then
                                    myDataRow = dtRTData.NewRow()
                                    myDataRow.ItemArray = DR(0).ItemArray
                                    myDataRow("AddFlag") = "1"
                                    myDataRow("SubInventory") = ""
                                    myDataRow("Locator") = ""
                                    myDataRow("CorrectType") = "TOTAL"
                                    myDataRow("QtyBaseUOM") = TotalQty

                                    dtRTData.Rows.Add(myDataRow)
                                End If
                            End If
                        Next
                        ds.Tables("RTData").Clear()

                        DR = Nothing
                        DR = dtRTData.Select("CorrectType <> ''", "MaterialNo, POLineID, CorrectType ")
                        For j = 0 To DR.Length - 1
                            ds.Tables("RTData").ImportRow(DR(j))
                            'ds.Tables("RTData").Rows(j).SetAdded()
                        Next
                    End If

                    Dim ReturnData As DataSet = New DataSet

                    Try
                        If ds.Tables(0).Rows.Count > 0 Then
                            ReturnData = Oracle_Return(LoginData, ds)
                        End If
                    Catch ex As Exception
                        ErrorLogging("Receiving-CancelGR1", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                        ReturnData = Nothing
                    End Try

                    If ReturnData Is Nothing OrElse ReturnData.Tables.Count = 0 OrElse ReturnData.Tables(0).Rows.Count = 0 Then
                        CancelGR.GRNo = ""
                        CancelGR.GRMessage = "Oracle Process Return Error"
                        CancelGR.GRStatus = "E"
                        Exit Function
                    End If


                    Dim ErrorTable As DataTable
                    Dim ErrorRow As Data.DataRow
                    ErrorTable = New Data.DataTable("ErrorTable")
                    ErrorTable.Columns.Add(New Data.DataColumn("ColumnName", System.Type.GetType("System.String")))
                    ErrorTable.Columns.Add(New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String")))


                    DR = Nothing
                    DR = ReturnData.Tables(0).Select("Status = 'N' or Status = ' '")
                    If DR.Length > 0 Then
                        For i = 0 To DR.Length - 1
                            If DR(i)("Message").ToString <> "" Then
                                ErrorRow = ErrorTable.NewRow()
                                ErrorRow("ColumnName") = DR(i)("MaterialNo")
                                ErrorRow("ErrorMsg") = DR(i)("Message").ToString
                                ErrorTable.Rows.Add(ErrorRow)
                            End If
                        Next
                        CancelGR.GRMsg = New DataSet
                        CancelGR.GRMsg.Tables.Add(ErrorTable)

                        CancelGR.GRNo = ""
                        CancelGR.GRMessage = "Oracle Process Return Error"
                        CancelGR.GRStatus = "E"
                    End If


                    DR = Nothing
                    DR = ReturnData.Tables(0).Select("Status = 'Y' ")
                    If DR.Length = 0 Then
                        CancelGR.GRNo = ""
                        CancelGR.GRMessage = "Oracle Process Return Error"
                        CancelGR.GRStatus = "E"
                        Exit Function
                    End If


                    DR = Nothing
                    DR = ReturnData.Tables(0).Select("Status = 'Y' and AddFlag = ''  ")
                    For i = 0 To DR.Length - 1
                        MaterialNo = DR(i)("MaterialNo").ToString
                        SubInventory = DR(i)("SubInventory").ToString
                        Locator = DR(i)("Locator").ToString
                        POLineID = DR(i)("POLineID").ToString

                        PORow = Nothing
                        PORow = Items.Tables(0).Select(" POLineID = '" & POLineID & "'")
                        If PORow.Length = 0 Then Continue For

                        CurrID = ""
                        PONo = PORow(0)("PONO")
                        POItem = PORow(0)("POItem")
                        ItemText = ReturnData.Tables(0).Rows(i)("RMANo")

                        CLIDRow = Nothing
                        'CLIDRow = Items.Tables(1).Select("Status = 'X' and PONO = '" & PONo & "' and POItem = '" & POItem & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
                        CLIDRow = Items.Tables(1).Select("Status = 'X' and POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
                        If CLIDRow.Length = 0 Then Continue For

                        Dim NoOfPackage As Integer = CLIDRow.Length
                        For j = 0 To CLIDRow.Length - 1
                            If CLIDRow(j)("RefCLID") Is DBNull.Value Then CLIDRow(j)("RefCLID") = ""
                            If CLIDRow(j)("RefCLID") <> "" Then
                                CLID = CLIDRow(j)("RefCLID")
                            Else
                                CLID = CLIDRow(j)("LID")
                            End If

                            If CurrID = CLID Then
                                NoOfPackage = NoOfPackage - 1
                                Continue For
                            End If
                            CurrID = CLID

                            Dim StartStr As String = Microsoft.VisualBasic.Left(CLID, 1)
                            Dim MidStr As String = Microsoft.VisualBasic.Mid(CLID, 3, 1)

                            Try
                                'Update CLMaster table
                                If StartStr = "B" OrElse StartStr = "P" OrElse MidStr = "V" Then           'Delete record from T_CLMaster if the LabelID is CartonID / PalletID, or Vendor CLID
                                    SqlStr = String.Format("DELETE from T_CLMaster Where CLID='{0}' ", CLID)
                                Else
                                    SqlStr = String.Format("Update T_CLMaster set StatusCode=0, ChangedOn=getdate(), LastTransaction='Receipt Reversal', ChangedBy='{0}', AddlText='{1}', BoxID='{2}' Where CLID='{3}' ", LoginData.User.ToUpper, ItemText, DBNull.Value, CLID)
                                End If
                                da.ExecuteNonQuery(SqlStr)
                            Catch ex As Exception
                                ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                            End Try


                            Try
                                'Update eJIT table
                                Dim eJITID As String
                                Dim JITQty, tmpQty As Decimal

                                SqlStr = String.Format("Select SourceID from T_CLMaster with (nolock) WHERE ProcessID='302' and CLID ='{0}'", CLID)
                                eJITID = Convert.ToString(da.ExecuteScalar(SqlStr))

                                If eJITID <> "" Then
                                    JITQty = CLIDRow(j)("Qty")

                                    SqlStr = String.Format("Select RecQty FROM T_KBPickingList with (nolock) WHERE EJITID='{0}'", eJITID)
                                    tmpQty = Convert.ToDecimal(da.ExecuteScalar(SqlStr))
                                    If tmpQty > 0 Then
                                        tmpQty = tmpQty - JITQty
                                        SqlStr = String.Format("UPDATE T_KBPickingList set RecQty='{1}' where EJITID='{0}'", eJITID, tmpQty)
                                        da.ExecuteNonQuery(SqlStr)
                                    End If
                                End If
                            Catch ex As Exception
                                ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                            End Try
                        Next


                        POType = DR(i)("POType").ToString

                        'Record the Pallet info which needs to do Reverse as well
                        'POType: Blank(Normal PO), OSP(OSP PO), K(KB PO), EJIT, IR, ASN
                        If POType = "" AndAlso RecConfig = "YES" Then
                            Dim DBdr() As DataRow = Nothing
                            DBdr = dtDashB.Select(" MaterialNo = '" & MaterialNo & "'")
                            If DBdr.Length = 0 Then
                                Dim myDR As DataRow
                                myDR = dtDashB.NewRow()
                                myDR("MaterialNo") = MaterialNo
                                myDR("Packages") = NoOfPackage
                                myDR("RecQty") = DR(i)("QtyBaseUOM")
                                dtDashB.Rows.Add(myDR)

                                'Record the MaterialNo Lists which have done reversal
                                If MatlLists = "" Then
                                    MatlLists = MaterialNo
                                Else
                                    MatlLists = MatlLists & "," & MaterialNo
                                End If
                            Else
                                DBdr(0)("Packages") = DBdr(0)("Packages") + NoOfPackage
                                DBdr(0)("RecQty") = DBdr(0)("RecQty") + DR(i)("QtyBaseUOM")
                            End If
                        End If
                    Next
                Catch ex As Exception
                    ErrorLogging("Receiving-CancelGR3", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
                End Try


                'Save the Pallet info in table T_RecDashboard which needs to do Component Delivery
                If RecConfig = "YES" AndAlso dtDashB.Rows.Count > 0 Then

                    MatlLists = "('" & MatlLists.Replace(",", "','") & "')"

                    ds = New DataSet
                    SqlStr = String.Format("Select MaterialNo,Packages,RecQty from T_RecDashboard with (nolock) WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo in {2}", LoginData.OrgCode, Header.OrderNo, MatlLists)
                    ds = da.ExecuteDataSet(SqlStr, "DBItems")

                    If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                    ElseIf ds.Tables(0).Rows.Count > 0 Then
                        For i = 0 To dtDashB.Rows.Count - 1
                            MaterialNo = dtDashB.Rows(i)("MaterialNo").ToString

                            DR = Nothing
                            DR = ds.Tables(0).Select(" MaterialNo = '" & MaterialNo & "'")
                            If DR.Length = 0 Then Continue For

                            Dim RestQty As Decimal = DR(0)("RecQty") - dtDashB.Rows(i)("RecQty")

                            If RestQty > 0 Then
                                Dim Remarks As String = DateTime.Now.ToString & " by User: " & LoginData.User.ToUpper & " with LastTransaction: Update RecQty by Receipt Reversal"
                                SqlStr = String.Format("Update T_RecDashboard set RecQty='{0}',Remarks='{1}' WHERE OrgCode='{2}' and RTNo ='{3}' and MaterialNo ='{4}'", RestQty, Remarks, LoginData.OrgCode, Header.OrderNo, MaterialNo)
                            Else
                                SqlStr = String.Format("Delete from T_RecDashboard WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo ='{2}'", LoginData.OrgCode, Header.OrderNo, MaterialNo)
                            End If
                            da.ExecuteNonQuery(SqlStr)
                        Next
                    End If
                End If


                ' Set Flag to show Partial complete or full complete
                If CancelGR.GRStatus Is Nothing Then
                    CancelGR.GRStatus = "Y"
                    CancelGR.GRMessage = "Reversal full completed "
                Else
                    CancelGR.GRStatus = "N"
                    CancelGR.GRMessage = "Reversal partial completed "
                End If
                CancelGR.GRNo = Header.OrderNo

                If Header.Type = 8 Then
                    CancelGR.CLIDs = New DataSet
                    CancelGR.CLIDs = Items
                End If

            Catch ex As Exception
                ErrorLogging("Receiving-CancelGR", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
            End Try

        End Using

    End Function

    Public Function Oracle_Return(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim aa As OracleString
            Dim oc As OracleCommand = da.Ora_Command_Trans()
            Dim oda As OracleDataAdapter = New OracleDataAdapter()

            Dim RTNo As String = p_ds.Tables("RTData").Rows(0)("RTNo").ToString
            Dim POType As String = p_ds.Tables("RTData").Rows(0)("POType").ToString

            Try

                Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables("RTData").Rows(0)("PONo"), 4, 2)
                If POStr = "04" Then LoginData.RespID_Inv = GetGlobalRespID(LoginData.Server)

                'For Org 680, there has Global PO which transferred from 580   --- 04/09/2019
                If LoginData.OrgCode = "680" AndAlso POStr <> "04" Then
                    Dim RTLot As String = p_ds.Tables("RTData").Rows(0)("RTLot").ToString
                    If RTLot <> "" AndAlso Not RTLot.Contains("CN") Then
                        LoginData.RespID_Inv = GetGlobalRespID(LoginData.Server)
                    End If
                End If

                oc.CommandType = CommandType.StoredProcedure
                oc.CommandText = "apps.xxetr_wip_pkg.initialize"
                oc.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
                oc.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
                oc.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID

                oc.ExecuteOracleNonQuery(aa)
                oc.Parameters.Clear()

                If POType = "ASN" Then             'Only ASN used PO Correction
                    oc.CommandText = "apps.xxetr_rcv_return_pkg.process_receiving_correct"                      'PO Correction for ASN
                    '        p_org_id            IN NUMBER,
                    '        p_transaction_id    IN NUMBER,
                    '        p_item_num          IN VARCHAR2,
                    '        p_quantity          IN NUMBER,
                    '        p_subinv            IN VARCHAR2,
                    '        p_locator           IN VARCHAR2,
                    '        p_lot_number        IN VARCHAR2,
                    '        p_correct_type      IN VARCHAR2,  --DELIVER / TOTAL
                    '        x_success_flag      OUT VARCHAR2,
                    '        x_error_mssg        OUT VARCHAR2
                    oc.Parameters.Add("p_org_id", OracleType.VarChar, 50).Value = LoginData.OrgID
                    oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
                    oc.Parameters.Add("p_item_num", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_quantity", OracleType.Double)
                    oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_lot_number", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_correct_type", OracleType.VarChar, 20)

                    oc.Parameters.Add("x_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                    oc.Parameters.Add("x_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                    oc.Parameters("p_transaction_id").SourceColumn = "POLineID"                      'TransactionID           -- 07/12/2017
                    oc.Parameters("p_item_num").SourceColumn = "MaterialNo"
                    oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
                    oc.Parameters("p_subinv").SourceColumn = "SubInventory"
                    oc.Parameters("p_locator").SourceColumn = "Locator"
                    oc.Parameters("p_lot_number").SourceColumn = "RTLot"
                    oc.Parameters("p_correct_type").SourceColumn = "CorrectType"

                    oc.Parameters("x_success_flag").SourceColumn = "Status"
                    oc.Parameters("x_error_mssg").SourceColumn = "Message"
                Else
                    If POType = "OSP" Then
                        oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return_osp"                             'OSP PO Reversal
                    Else
                        oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return"                                      'Normal PO / KanBan PO Reversal
                    End If

                    oc.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = LoginData.OrgCode
                    oc.Parameters.Add("p_po_number", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_receipt_num", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
                    oc.Parameters.Add("p_line_num", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_shipment_num", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_release_num", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_quantity", OracleType.Double)
                    oc.Parameters.Add("p_uom_code", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
                    oc.Parameters.Add("p_reason_code", OracleType.VarChar, 200).Value = ""    'Use default ReasonCode from Oracle
                    oc.Parameters.Add("p_rma_number", OracleType.VarChar, 500)

                    oc.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
                    oc.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                    oc.Parameters("p_po_number").SourceColumn = "PONo"
                    oc.Parameters("p_receipt_num").SourceColumn = "RTNo"
                    oc.Parameters("p_transaction_id").SourceColumn = "TranID"                      'TransactionID           -- 07/12/2017
                    oc.Parameters("p_line_num").SourceColumn = "POLine"
                    oc.Parameters("p_shipment_num").SourceColumn = "ShipmentNo"
                    oc.Parameters("p_release_num").SourceColumn = "ReleaseNo"
                    oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
                    oc.Parameters("p_uom_code").SourceColumn = "BaseUOM"
                    oc.Parameters("p_subinv").SourceColumn = "SubInventory"
                    oc.Parameters("p_locator").SourceColumn = "Locator"
                    oc.Parameters("p_rma_number").SourceColumn = "RMANo"
                    oc.Parameters("o_success_flag").SourceColumn = "Status"
                    oc.Parameters("o_error_mssg").SourceColumn = "Message"
                End If

                oda.InsertCommand = oc
                oda.Update(p_ds.Tables("RTData"))

                Dim DR() As DataRow = Nothing
                DR = p_ds.Tables("RTData").Select("Status = 'N' or Status = ' '")
                If DR.Length = 0 Then
                    oc.Transaction.Commit()
                    oc.Connection.Close()
                    'oc.Connection.Dispose()
                    'oc.Dispose()
                    Return p_ds
                    Exit Function
                End If


                'Record error message to eTrace database
                Dim i As Integer
                Dim ErrMsg As String = ""
                For i = 0 To DR.Length - 1
                    If DR(i)("Message").ToString.Trim <> "" Then
                        If ErrMsg = "" Then
                            ErrMsg = DR(i)("Message").ToString.Trim
                        Else
                            ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString.Trim
                        End If
                    End If
                Next
                ErrorLogging("Receiving-Oracle_Return1", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with return message: " & ErrMsg, "I")

                oc.Transaction.Rollback()
                oc.Connection.Close()
                'oc.Connection.Dispose()
                'oc.Dispose()

                Return p_ds

            Catch oe As Exception
                ErrorLogging("Receiving-Oracle_Return", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with error message: " & oe.Message & oe.Source, "E")
                p_ds.Tables("RTData").Rows(0)("Status") = "N"
                p_ds.Tables("RTData").Rows(0)("Message") = oe.Message
                p_ds.Tables("RTData").Rows(0).AcceptChanges()

                If oc.Connection.State = ConnectionState.Open Then
                    oc.Transaction.Rollback()
                    oc.Connection.Close()
                End If
                Return p_ds
            Finally
                If oc.Connection.State <> ConnectionState.Closed Then oc.Connection.Close()
            End Try

        End Using
    End Function

    'Private Function CancelGR(ByVal LoginData As ERPLogin, ByVal Header As GRHeaderStructure, ByVal Items As DataSet) As CreateGRResponse
    '    Using da As DataAccess = GetDataAccess()

    '        Try

    '            Dim ds As DataSet = New DataSet
    '            Dim RTData As DataTable
    '            Dim myDataRow As Data.DataRow

    '            RTData = New Data.DataTable("RTData")
    '            RTData.Columns.Add(New Data.DataColumn("RTNo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("POLineID", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("PONo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("ReleaseNo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("POLine", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("ShipmentNo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("QtyBaseUOM", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("BaseUOM", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("SubInventory", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("Locator", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("RTLot", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("RMANo", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("POType", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("AddFlag", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("CorrectType", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("Status", System.Type.GetType("System.String")))
    '            RTData.Columns.Add(New Data.DataColumn("Message", System.Type.GetType("System.String")))
    '            ds.Tables.Add(RTData)

    '            Dim dtDashB As DataTable = New DataTable
    '            dtDashB.Columns.Add(New Data.DataColumn("MaterialNo", System.Type.GetType("System.String")))
    '            dtDashB.Columns.Add(New Data.DataColumn("Packages", System.Type.GetType("System.String")))
    '            dtDashB.Columns.Add(New Data.DataColumn("RecQty", System.Type.GetType("System.Decimal")))

    '            Dim dtDelivery As DataTable = New DataTable
    '            dtDelivery.Columns.Add(New Data.DataColumn("POLineID", System.Type.GetType("System.String")))
    '            dtDelivery.Columns.Add(New Data.DataColumn("Qty", System.Type.GetType("System.Decimal")))

    '            Dim i, j As Integer
    '            Dim DR(), PORow(), CLIDRow() As DataRow
    '            Dim PONo, POItem, ItemText, CLID, CurrID As String
    '            Dim POLineID, MaterialNo, SubInventory, Locator As String

    '            Dim SqlStr As String
    '            Dim MatlLists As String = ""
    '            Dim RecConfig As String = ""

    '            Try
    '                'Read Receiving DashBoard Flag from Table T_Config to decide if need to Delete PalletID info from table T_RecDashboard 
    '                SqlStr = String.Format("Select Value from T_Config with (nolock) where ConfigID = 'REC001'")
    '                RecConfig = Convert.ToString(da.ExecuteScalar(SqlStr))

    '                Dim CLIDTable As DataTable = New DataTable
    '                CLIDTable = Items.Tables(1).Copy

    '                CLIDRow = Nothing
    '                CLIDRow = CLIDTable.Select(" Status = 'X' and RefCLID IS NOT NULL and RefCLID <> ' ' ")
    '                If CLIDRow.Length > 0 Then
    '                    For i = 0 To CLIDRow.Length - 1
    '                        Dim RefCLID As String = CLIDRow(i)("RefCLID").ToString
    '                        DR = Nothing
    '                        DR = Items.Tables(1).Select(" RefCLID = '" & RefCLID & "'")
    '                        If DR.Length > 0 Then
    '                            For j = 0 To DR.Length - 1
    '                                If DR(j)("Status").ToString = "" Then
    '                                    DR(j)("Status") = "X"
    '                                    DR(j).AcceptChanges()
    '                                End If
    '                            Next
    '                            Items.Tables(1).AcceptChanges()
    '                        End If
    '                    Next
    '                End If

    '                Dim POType As String = Items.Tables(0).Rows(0)("PullMethod").ToString

    '                For i = 0 To Items.Tables(1).Rows.Count - 1
    '                    If Items.Tables(1).Rows(i)("Status") Is DBNull.Value OrElse Items.Tables(1).Rows(i)("Status") = "" Then
    '                    ElseIf Items.Tables(1).Rows(i)("Status") = "X" Then

    '                        PONo = Items.Tables(1).Rows(i)("PONO").ToString
    '                        POItem = Items.Tables(1).Rows(i)("POItem").ToString
    '                        POLineID = Items.Tables(1).Rows(i)("POLineID").ToString
    '                        SubInventory = Items.Tables(1).Rows(i)("SubInventory").ToString
    '                        Locator = Items.Tables(1).Rows(i)("Locator").ToString

    '                        'For Non OSP PO Receipt, Collect Qty by TransactionID here for later Correction
    '                        If POType <> "OSP" Then
    '                            PORow = Nothing
    '                            PORow = dtDelivery.Select(" POLineID = '" & POLineID & "'")
    '                            If PORow.Length = 0 Then
    '                                myDataRow = dtDelivery.NewRow()
    '                                myDataRow("POLineID") = POLineID
    '                                myDataRow("Qty") = Convert.ToDecimal(Items.Tables(1).Rows(i)("Qty"))
    '                                dtDelivery.Rows.Add(myDataRow)
    '                            Else
    '                                Dim Qty As Decimal = PORow(0)("Qty")
    '                                PORow(0)("Qty") = Qty + Items.Tables(1).Rows(i)("Qty")
    '                                PORow(0).AcceptChanges()
    '                            End If

    '                        End If

    '                        PORow = Nothing
    '                        PORow = Items.Tables(0).Select(" PONO = '" & PONo & "' and POItem = '" & POItem & "' and POLineID = '" & POLineID & "'")
    '                        If PORow.Length > 0 Then
    '                            DR = Nothing
    '                            DR = ds.Tables("RTData").Select(" POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                            If DR.Length = 0 Then
    '                                myDataRow = ds.Tables("RTData").NewRow()
    '                                myDataRow("RTNo") = Header.OrderNo
    '                                myDataRow("MaterialNo") = Items.Tables(1).Rows(i)("MatNo")
    '                                myDataRow("POLineID") = POLineID

    '                                Dim OrderData As POData
    '                                OrderData = Split_POData(PONo, POItem)

    '                                myDataRow("PONo") = OrderData.PONo
    '                                myDataRow("POLine") = OrderData.LineNo
    '                                myDataRow("ShipmentNo") = OrderData.ShipmentNo
    '                                If OrderData.ReleaseNo = 0 Then
    '                                    myDataRow("ReleaseNo") = ""
    '                                Else
    '                                    myDataRow("ReleaseNo") = OrderData.ReleaseNo
    '                                End If

    '                                myDataRow("QtyBaseUOM") = Convert.ToDecimal(Items.Tables(1).Rows(i)("Qty"))
    '                                myDataRow("BaseUOM") = PORow(0)("BaseUOM").ToString
    '                                myDataRow("SubInventory") = SubInventory
    '                                myDataRow("Locator") = Locator
    '                                myDataRow("RTLot") = PORow(0)("RTLot").ToString
    '                                myDataRow("RMANo") = PORow(0)("ItemText").ToString
    '                                myDataRow("POType") = PORow(0)("PullMethod").ToString
    '                                myDataRow("AddFlag") = ""
    '                                myDataRow("Status") = ""
    '                                myDataRow("Message") = ""
    '                                ds.Tables("RTData").Rows.Add(myDataRow)
    '                            Else
    '                                Dim QtyBaseUOM As Decimal
    '                                QtyBaseUOM = DR(0)("QtyBaseUOM")

    '                                DR(0)("QtyBaseUOM") = QtyBaseUOM + Items.Tables(1).Rows(i)("Qty")
    '                                DR(0).AcceptChanges()
    '                                DR(0).SetAdded()
    '                            End If

    '                            Items.Tables(0).AcceptChanges()
    '                        End If
    '                    End If
    '                Next

    '                'For Non OSP PO Receipt, Collect Qty by TransactionID here for later Correction
    '                If POType <> "OSP" Then
    '                    Dim dtRTData As New DataTable
    '                    dtRTData = ds.Tables("RTData").Clone

    '                    For i = 0 To dtDelivery.Rows.Count - 1
    '                        Dim TotalFlag As Boolean = False
    '                        Dim TotalQty As Decimal = dtDelivery.Rows(i)("Qty")

    '                        DR = Nothing
    '                        DR = ds.Tables("RTData").Select(" POLineID = '" & dtDelivery.Rows(i)("POLineID") & "'")
    '                        If DR.Length > 0 Then
    '                            For j = 0 To DR.Length - 1
    '                                Dim mySubInv As String = DR(j)("SubInventory").ToString
    '                                If mySubInv = "" Then
    '                                    TotalFlag = True
    '                                    DR(j)("CorrectType") = "TOTAL"
    '                                    DR(j)("QtyBaseUOM") = TotalQty
    '                                    dtRTData.ImportRow(DR(j))
    '                                Else
    '                                    DR(j)("CorrectType") = "DELIVER"
    '                                    dtRTData.ImportRow(DR(j))
    '                                End If
    '                            Next

    '                            'If All Qty has done Delivery, need to add one row which has Total Qty
    '                            If TotalFlag = False Then
    '                                myDataRow = dtRTData.NewRow()
    '                                myDataRow.ItemArray = DR(0).ItemArray
    '                                myDataRow("AddFlag") = "1"
    '                                myDataRow("SubInventory") = ""
    '                                myDataRow("Locator") = ""
    '                                myDataRow("CorrectType") = "TOTAL"
    '                                myDataRow("QtyBaseUOM") = TotalQty

    '                                dtRTData.Rows.Add(myDataRow)
    '                            End If
    '                        End If
    '                    Next
    '                    ds.Tables("RTData").Clear()

    '                    DR = Nothing
    '                    DR = dtRTData.Select("CorrectType <> ''", "MaterialNo, POLineID, CorrectType ")
    '                    For j = 0 To DR.Length - 1
    '                        ds.Tables("RTData").ImportRow(DR(j))
    '                        'ds.Tables("RTData").Rows(j).SetAdded()
    '                    Next
    '                End If

    '                Dim ReturnData As DataSet = New DataSet

    '                Try
    '                    If ds.Tables(0).Rows.Count > 0 Then
    '                        ReturnData = Oracle_Correct(LoginData, ds)
    '                    End If
    '                Catch ex As Exception
    '                    ErrorLogging("Receiving-CancelGR1", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                    ReturnData = Nothing
    '                End Try

    '                If ReturnData Is Nothing OrElse ReturnData.Tables.Count = 0 OrElse ReturnData.Tables(0).Rows.Count = 0 Then
    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                    Exit Function
    '                End If


    '                Dim ErrorTable As DataTable
    '                Dim ErrorRow As Data.DataRow
    '                ErrorTable = New Data.DataTable("ErrorTable")
    '                ErrorTable.Columns.Add(New Data.DataColumn("ColumnName", System.Type.GetType("System.String")))
    '                ErrorTable.Columns.Add(New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String")))


    '                DR = Nothing
    '                DR = ReturnData.Tables(0).Select("Status = 'N' or Status = ' '")
    '                If DR.Length > 0 Then
    '                    For i = 0 To DR.Length - 1
    '                        If DR(i)("Message").ToString <> "" Then
    '                            ErrorRow = ErrorTable.NewRow()
    '                            ErrorRow("ColumnName") = DR(i)("MaterialNo")
    '                            ErrorRow("ErrorMsg") = DR(i)("Message").ToString
    '                            ErrorTable.Rows.Add(ErrorRow)
    '                        End If
    '                    Next
    '                    CancelGR.GRMsg = New DataSet
    '                    CancelGR.GRMsg.Tables.Add(ErrorTable)

    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                End If


    '                DR = Nothing
    '                DR = ReturnData.Tables(0).Select("Status = 'Y' ")
    '                If DR.Length = 0 Then
    '                    CancelGR.GRNo = ""
    '                    CancelGR.GRMessage = "Oracle Process Return Error"
    '                    CancelGR.GRStatus = "E"
    '                    Exit Function
    '                End If


    '                DR = Nothing
    '                DR = ReturnData.Tables(0).Select("Status = 'Y' and AddFlag = ''  ")
    '                For i = 0 To DR.Length - 1
    '                    MaterialNo = DR(i)("MaterialNo").ToString
    '                    SubInventory = DR(i)("SubInventory").ToString
    '                    Locator = DR(i)("Locator").ToString
    '                    POLineID = DR(i)("POLineID").ToString

    '                    PORow = Nothing
    '                    PORow = Items.Tables(0).Select(" POLineID = '" & POLineID & "'")
    '                    If PORow.Length = 0 Then Continue For

    '                    CurrID = ""
    '                    PONo = PORow(0)("PONO")
    '                    POItem = PORow(0)("POItem")
    '                    ItemText = ReturnData.Tables(0).Rows(i)("RMANo")

    '                    CLIDRow = Nothing
    '                    'CLIDRow = Items.Tables(1).Select("Status = 'X' and PONO = '" & PONo & "' and POItem = '" & POItem & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                    CLIDRow = Items.Tables(1).Select("Status = 'X' and POLineID = '" & POLineID & "' and SubInventory = '" & SubInventory & "' and Locator = '" & Locator & "'")
    '                    If CLIDRow.Length = 0 Then Continue For

    '                    Dim NoOfPackage As Integer = CLIDRow.Length
    '                    For j = 0 To CLIDRow.Length - 1
    '                        If CLIDRow(j)("RefCLID") Is DBNull.Value Then CLIDRow(j)("RefCLID") = ""
    '                        If CLIDRow(j)("RefCLID") <> "" Then
    '                            CLID = CLIDRow(j)("RefCLID")
    '                        Else
    '                            CLID = CLIDRow(j)("LID")
    '                        End If

    '                        If CurrID = CLID Then
    '                            NoOfPackage = NoOfPackage - 1
    '                            Continue For
    '                        End If
    '                        CurrID = CLID

    '                        Dim StartStr As String = Microsoft.VisualBasic.Left(CLID, 1)
    '                        Dim MidStr As String = Microsoft.VisualBasic.Mid(CLID, 3, 1)

    '                        Try
    '                            'Update CLMaster table
    '                            If StartStr = "B" OrElse StartStr = "P" OrElse MidStr = "V" Then           'Delete record from T_CLMaster if the LabelID is CartonID / PalletID, or Vendor CLID
    '                                SqlStr = String.Format("DELETE from T_CLMaster Where CLID='{0}' ", CLID)
    '                            Else
    '                                SqlStr = String.Format("Update T_CLMaster set StatusCode=0, ChangedOn=getdate(), LastTransaction='Receipt Reversal', ChangedBy='{0}', AddlText='{1}', BoxID='{2}' Where CLID='{3}' ", LoginData.User.ToUpper, ItemText, DBNull.Value, CLID)
    '                            End If
    '                            da.ExecuteNonQuery(SqlStr)
    '                        Catch ex As Exception
    '                            ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                        End Try


    '                        Try
    '                            'Update eJIT table
    '                            Dim eJITID As String
    '                            Dim JITQty, tmpQty As Decimal

    '                            SqlStr = String.Format("Select SourceID from T_CLMaster with (nolock) WHERE ProcessID='302' and CLID ='{0}'", CLID)
    '                            eJITID = Convert.ToString(da.ExecuteScalar(SqlStr))

    '                            If eJITID <> "" Then
    '                                JITQty = CLIDRow(j)("Qty")

    '                                SqlStr = String.Format("Select RecQty FROM T_KBPickingList with (nolock) WHERE EJITID='{0}'", eJITID)
    '                                tmpQty = Convert.ToDecimal(da.ExecuteScalar(SqlStr))
    '                                If tmpQty > 0 Then
    '                                    tmpQty = tmpQty - JITQty
    '                                    SqlStr = String.Format("UPDATE T_KBPickingList set RecQty='{1}' where EJITID='{0}'", eJITID, tmpQty)
    '                                    da.ExecuteNonQuery(SqlStr)
    '                                End If
    '                            End If
    '                        Catch ex As Exception
    '                            ErrorLogging("Receiving-CancelGR2", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '                        End Try
    '                    Next


    '                    POType = DR(i)("POType").ToString

    '                    'Record the Pallet info which needs to do Reverse as well
    '                    'POType: Blank(Normal PO), OSP(OSP PO), K(KB PO), EJIT, IR
    '                    If POType = "" AndAlso RecConfig = "YES" Then
    '                        Dim DBdr() As DataRow = Nothing
    '                        DBdr = dtDashB.Select(" MaterialNo = '" & MaterialNo & "'")
    '                        If DBdr.Length = 0 Then
    '                            Dim myDR As DataRow
    '                            myDR = dtDashB.NewRow()
    '                            myDR("MaterialNo") = MaterialNo
    '                            myDR("Packages") = NoOfPackage
    '                            myDR("RecQty") = DR(i)("QtyBaseUOM")
    '                            dtDashB.Rows.Add(myDR)

    '                            'Record the MaterialNo Lists which have done reversal
    '                            If MatlLists = "" Then
    '                                MatlLists = MaterialNo
    '                            Else
    '                                MatlLists = MatlLists & "," & MaterialNo
    '                            End If
    '                        Else
    '                            DBdr(0)("Packages") = DBdr(0)("Packages") + NoOfPackage
    '                            DBdr(0)("RecQty") = DBdr(0)("RecQty") + DR(i)("QtyBaseUOM")
    '                        End If
    '                    End If
    '                Next
    '            Catch ex As Exception
    '                ErrorLogging("Receiving-CancelGR3", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '            End Try


    '            'Save the Pallet info in table T_RecDashboard which needs to do Component Delivery
    '            If RecConfig = "YES" AndAlso dtDashB.Rows.Count > 0 Then

    '                MatlLists = "('" & MatlLists.Replace(",", "','") & "')"

    '                ds = New DataSet
    '                SqlStr = String.Format("Select MaterialNo,Packages,RecQty from T_RecDashboard with (nolock) WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo in {2}", LoginData.OrgCode, Header.OrderNo, MatlLists)
    '                ds = da.ExecuteDataSet(SqlStr, "DBItems")

    '                If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
    '                ElseIf ds.Tables(0).Rows.Count > 0 Then
    '                    For i = 0 To dtDashB.Rows.Count - 1
    '                        MaterialNo = dtDashB.Rows(i)("MaterialNo").ToString

    '                        DR = Nothing
    '                        DR = ds.Tables(0).Select(" MaterialNo = '" & MaterialNo & "'")
    '                        If DR.Length = 0 Then Continue For

    '                        Dim RestQty As Decimal = DR(0)("RecQty") - dtDashB.Rows(i)("RecQty")

    '                        If RestQty > 0 Then
    '                            Dim Remarks As String = DateTime.Now.ToString & " by User: " & LoginData.User.ToUpper & " with LastTransaction: Update RecQty by Receipt Reversal"
    '                            SqlStr = String.Format("Update T_RecDashboard set RecQty='{0}',Remarks='{1}' WHERE OrgCode='{2}' and RTNo ='{3}' and MaterialNo ='{4}'", RestQty, Remarks, LoginData.OrgCode, Header.OrderNo, MaterialNo)
    '                        Else
    '                            SqlStr = String.Format("Delete from T_RecDashboard WHERE OrgCode='{0}' and RTNo ='{1}' and MaterialNo ='{2}'", LoginData.OrgCode, Header.OrderNo, MaterialNo)
    '                        End If
    '                        da.ExecuteNonQuery(SqlStr)
    '                    Next
    '                End If
    '            End If


    '            ' Set Flag to show Partial complete or full complete
    '            If CancelGR.GRStatus Is Nothing Then
    '                CancelGR.GRStatus = "Y"
    '                CancelGR.GRMessage = "Reversal full completed "
    '            Else
    '                CancelGR.GRStatus = "N"
    '                CancelGR.GRMessage = "Reversal partial completed "
    '            End If
    '            CancelGR.GRNo = Header.OrderNo

    '            If Header.Type = 8 Then
    '                CancelGR.CLIDs = New DataSet
    '                CancelGR.CLIDs = Items
    '            End If

    '        Catch ex As Exception
    '            ErrorLogging("Receiving-CancelGR", LoginData.User.ToUpper, "ReceiptNo: " & Header.OrderNo & " with Receipt type " & Header.Type & ", " & ex.Message & ex.Source, "E")
    '        End Try

    '    End Using

    'End Function

    'Public Function Oracle_Correct(ByVal LoginData As ERPLogin, ByVal p_ds As DataSet) As DataSet
    '    Using da As DataAccess = GetDataAccess()

    '        Dim aa As OracleString
    '        Dim oc As OracleCommand = da.Ora_Command_Trans()
    '        Dim oda As OracleDataAdapter = New OracleDataAdapter()

    '        Dim RTNo As String = p_ds.Tables("RTData").Rows(0)("RTNo").ToString
    '        Dim POType As String = p_ds.Tables("RTData").Rows(0)("POType").ToString

    '        Try

    '            Dim POStr As String = Microsoft.VisualBasic.Mid(p_ds.Tables("RTData").Rows(0)("PONo"), 4, 2)
    '            If POStr = "04" Then LoginData.RespID_Inv = GetGlobalRespID(LoginData.Server)


    '            oc.CommandType = CommandType.StoredProcedure
    '            oc.CommandText = "apps.xxetr_wip_pkg.initialize"
    '            oc.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(LoginData.UserID)
    '            oc.Parameters.Add("p_resp_id", OracleType.Int32).Value = LoginData.RespID_Inv    'RespID
    '            oc.Parameters.Add("p_appl_id", OracleType.Int32).Value = LoginData.AppID_Inv     'AppID

    '            oc.ExecuteOracleNonQuery(aa)
    '            oc.Parameters.Clear()

    '            If POType = "OSP" Then
    '                oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return_osp"       'OSP PO Reversal
    '                oc.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = LoginData.OrgCode
    '                oc.Parameters.Add("p_po_number", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_receipt_num", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
    '                oc.Parameters.Add("p_line_num", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_shipment_num", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_release_num", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_quantity", OracleType.Double)
    '                oc.Parameters.Add("p_uom_code", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_reason_code", OracleType.VarChar, 200).Value = ""    'Use default ReasonCode from Oracle
    '                oc.Parameters.Add("p_rma_number", OracleType.VarChar, 500)

    '                oc.Parameters.Add("o_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
    '                oc.Parameters.Add("o_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

    '                oc.Parameters("p_po_number").SourceColumn = "PONo"
    '                oc.Parameters("p_receipt_num").SourceColumn = "RTNo"
    '                oc.Parameters("p_transaction_id").SourceColumn = "POLineID"                      'TransactionID           -- 07/12/2017
    '                oc.Parameters("p_line_num").SourceColumn = "POLine"
    '                oc.Parameters("p_shipment_num").SourceColumn = "ShipmentNo"
    '                oc.Parameters("p_release_num").SourceColumn = "ReleaseNo"
    '                oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
    '                oc.Parameters("p_uom_code").SourceColumn = "BaseUOM"
    '                oc.Parameters("p_subinv").SourceColumn = "SubInventory"
    '                oc.Parameters("p_locator").SourceColumn = "Locator"
    '                oc.Parameters("p_rma_number").SourceColumn = "RMANo"
    '                oc.Parameters("o_success_flag").SourceColumn = "Status"
    '                oc.Parameters("o_error_mssg").SourceColumn = "Message"
    '            Else
    '                'oc.CommandText = "apps.xxetr_rcv_return_pkg.process_return"                                 'Normal PO / KanBan PO Reversal
    '                oc.CommandText = "apps.xxetr_rcv_return_pkg.process_receiving_correct"               'Normal PO Correction
    '                '        p_org_id            IN NUMBER,
    '                '        p_transaction_id    IN NUMBER,
    '                '        p_item_num          IN VARCHAR2,
    '                '        p_quantity          IN NUMBER,
    '                '        p_subinv            IN VARCHAR2,
    '                '        p_locator           IN VARCHAR2,
    '                '        p_lot_number        IN VARCHAR2,
    '                '        p_correct_type      IN VARCHAR2,  --DELIVER / TOTAL
    '                '        x_success_flag      OUT VARCHAR2,
    '                '        x_error_mssg        OUT VARCHAR2
    '                oc.Parameters.Add("p_org_id", OracleType.VarChar, 50).Value = LoginData.OrgID
    '                oc.Parameters.Add("p_transaction_id", OracleType.Int32)                              'p_transaction_id
    '                oc.Parameters.Add("p_item_num", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_quantity", OracleType.Double)
    '                oc.Parameters.Add("p_subinv", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_locator", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_lot_number", OracleType.VarChar, 50)
    '                oc.Parameters.Add("p_correct_type", OracleType.VarChar, 20)

    '                oc.Parameters.Add("x_success_flag", OracleType.VarChar, 10).Direction = ParameterDirection.Output
    '                oc.Parameters.Add("x_error_mssg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

    '                oc.Parameters("p_transaction_id").SourceColumn = "POLineID"                      'TransactionID           -- 07/12/2017
    '                oc.Parameters("p_item_num").SourceColumn = "MaterialNo"
    '                oc.Parameters("p_quantity").SourceColumn = "QtyBaseUOM"
    '                oc.Parameters("p_subinv").SourceColumn = "SubInventory"
    '                oc.Parameters("p_locator").SourceColumn = "Locator"
    '                oc.Parameters("p_lot_number").SourceColumn = "RTLot"
    '                oc.Parameters("p_correct_type").SourceColumn = "CorrectType"

    '                oc.Parameters("x_success_flag").SourceColumn = "Status"
    '                oc.Parameters("x_error_mssg").SourceColumn = "Message"
    '            End If

    '            oda.InsertCommand = oc
    '            oda.Update(p_ds.Tables("RTData"))

    '            Dim DR() As DataRow = Nothing
    '            DR = p_ds.Tables("RTData").Select("Status = 'N' or Status = ' '")
    '            If DR.Length = 0 Then
    '                oc.Transaction.Commit()
    '                oc.Connection.Close()
    '                'oc.Connection.Dispose()
    '                'oc.Dispose()
    '                Return p_ds
    '                Exit Function
    '            End If


    '            'Record error message to eTrace database
    '            Dim i As Integer
    '            Dim ErrMsg As String = ""
    '            For i = 0 To DR.Length - 1
    '                If DR(i)("Message").ToString.Trim <> "" Then
    '                    If ErrMsg = "" Then
    '                        ErrMsg = DR(i)("Message").ToString.Trim
    '                    Else
    '                        ErrMsg = ErrMsg & "; " & DR(i)("Message").ToString.Trim
    '                    End If
    '                End If
    '            Next
    '            ErrorLogging("Receiving-Oracle_Correct1", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with return message: " & ErrMsg, "I")

    '            oc.Transaction.Rollback()
    '            oc.Connection.Close()
    '            'oc.Connection.Dispose()
    '            'oc.Dispose()

    '            Return p_ds

    '        Catch oe As Exception
    '            ErrorLogging("Receiving-Oracle_Correct", LoginData.User.ToUpper, "ReceiptNo: " & RTNo & " with error message: " & oe.Message & oe.Source, "E")
    '            p_ds.Tables("RTData").Rows(0)("Status") = "N"
    '            p_ds.Tables("RTData").Rows(0)("Message") = oe.Message
    '            p_ds.Tables("RTData").Rows(0).AcceptChanges()

    '            If oc.Connection.State = ConnectionState.Open Then
    '                oc.Transaction.Rollback()
    '                oc.Connection.Close()
    '            End If
    '            Return p_ds
    '        Finally
    '            If oc.Connection.State <> ConnectionState.Closed Then oc.Connection.Close()
    '        End Try

    '    End Using
    'End Function

    Public Function Split_POData(ByVal OrderNo As String, ByVal OrderItem As String) As POData
        Dim OrderData As POData
        Dim POArry(), LineArry() As String

        Try
            If OrderNo Is Nothing OrElse OrderNo = "" Then
                OrderData.PONo = ""
                OrderData.ReleaseNo = 0
            ElseIf OrderNo.ToString <> "" Then
                If IsNumeric(OrderNo) Then
                    OrderData.PONo = OrderNo
                    OrderData.ReleaseNo = 0
                Else
                    POArry = Split(OrderNo, "-")
                    If POArry.Length = 2 Then
                        OrderData.PONo = POArry(0).ToString
                        OrderData.ReleaseNo = POArry(1).ToString
                    Else
                        OrderData.PONo = POArry(0).ToString
                        OrderData.ReleaseNo = 0
                    End If
                End If
            End If

            If OrderItem Is Nothing OrElse OrderItem = "" Then
                OrderData.LineNo = 0
                OrderData.ShipmentNo = 0
                OrderData.Distribution = 0
                OrderData.MaterialNo = ""
                OrderData.Locator = ""

            ElseIf OrderItem.ToString <> "" Then
                If OrderItem.ToString.StartsWith("P") Then
                    OrderData.LineNo = 0
                    OrderData.ShipmentNo = 0
                    OrderData.Distribution = 0
                    OrderData.MaterialNo = Mid(OrderItem, 2)
                    OrderData.Locator = ""

                    ' Get Locator 
                ElseIf OrderItem.ToString.StartsWith("L") Then
                    OrderData.LineNo = 0
                    OrderData.ShipmentNo = 0
                    OrderData.Distribution = 0
                    OrderData.Locator = Mid(OrderItem, 2)
                    OrderData.MaterialNo = ""
                Else
                    OrderData.MaterialNo = ""
                    OrderData.Locator = ""

                    LineArry = Split(OrderItem, ".")
                    If LineArry.Length = 3 Then    'Distribution
                        OrderData.LineNo = LineArry(0).ToString
                        OrderData.ShipmentNo = LineArry(1).ToString
                        OrderData.Distribution = LineArry(2).ToString
                    ElseIf LineArry.Length = 2 Then
                        OrderData.LineNo = LineArry(0).ToString
                        OrderData.ShipmentNo = LineArry(1).ToString
                        OrderData.Distribution = 0
                    Else
                        OrderData.LineNo = LineArry(0).ToString
                        OrderData.ShipmentNo = 0
                        OrderData.Distribution = 0
                    End If
                End If

            End If

            Return OrderData

        Catch ex As Exception
            OrderData.PONo = ""
            Return OrderData
        End Try

    End Function

    Public Function ReadShipmentData(ByVal LoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs = New DataSet

            Dim OrderLists As String = " ( '" & OrderNo.Replace(",", "','") & "') "

            Try
                Dim Sqlstr1, Sqlstr2 As String
                Sqlstr1 = String.Format("Select CLID, QtyBaseUOM, BoxID, PurOrdNo as ProdOrder from T_CLMaster with (nolock) where ProcessID = 2 and PurOrdNo in {0} and OrgCode = '{1}'", OrderLists, LoginData.OrgCode)
                Sqlstr2 = String.Format("Select CartonID as CLID, count ( CartonID ) as QtyBaseUOM, PalletID as BoxID, ProdOrder from T_Shippment with (nolock) where StatusCode = 1 and ProdOrder in {0} GROUP BY CartonID, PalletID, ProdOrder ", OrderLists)

                Dim sql() As String = {Sqlstr1, Sqlstr2}
                Dim table() As String = {"CLIDData", "ShipData"}

                myCLIDs = da.ExecuteDataSet(sql, table)

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                    Return Nothing
                    Exit Function
                End If

                Dim i, j As Integer
                Dim DR() As DataRow = Nothing
                If myCLIDs.Tables("CLIDData").Rows.Count > 0 Then
                    For i = 0 To myCLIDs.Tables("CLIDData").Rows.Count - 1
                        Dim LabelID As String = myCLIDs.Tables("CLIDData").Rows(i)("CLID")
                        DR = myCLIDs.Tables("ShipData").Select(" CLID = '" & LabelID & "'")
                        If DR.Length > 0 Then
                            For j = 0 To DR.Length - 1
                                DR(j).Delete()
                            Next
                            myCLIDs.Tables("ShipData").AcceptChanges()
                        End If
                    Next
                    myCLIDs.AcceptChanges()
                End If

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("Receiving-ReadShipmentData", LoginData.User.ToUpper, "OSPDJ: " & OrderNo & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function TestShipmentData(ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim LoginData As ERPLogin = New ERPLogin
            LoginData.OrgCode = "409"

            Dim myCLIDs = New DataSet

            Dim OrderLists As String = " ( '" & OrderNo.Replace(",", "','") & "') "

            Try
                Dim Sqlstr1, Sqlstr2 As String
                Sqlstr1 = String.Format("Select CLID, QtyBaseUOM, BoxID, PurOrdNo as ProdOrder from T_CLMaster with (nolock) where ProcessID = 2 and PurOrdNo in {0} and OrgCode = '{1}'", OrderLists, LoginData.OrgCode)
                Sqlstr2 = String.Format("Select CartonID as CLID, count ( CartonID ) as QtyBaseUOM, PalletID as BoxID, ProdOrder from T_Shippment with (nolock) where StatusCode = 1 and ProdOrder in {0} GROUP BY CartonID, PalletID, ProdOrder ", OrderLists)

                Dim sql() As String = {Sqlstr1, Sqlstr2}
                Dim table() As String = {"CLIDData", "ShipData"}

                myCLIDs = da.ExecuteDataSet(sql, table)

                If myCLIDs Is Nothing OrElse myCLIDs.Tables.Count = 0 Then
                    Return Nothing
                    Exit Function
                End If

                Dim i, j As Integer
                Dim DR() As DataRow = Nothing
                If myCLIDs.Tables("CLIDData").Rows.Count > 0 Then
                    For i = 0 To myCLIDs.Tables("CLIDData").Rows.Count - 1
                        Dim LabelID As String = myCLIDs.Tables("CLIDData").Rows(i)("CLID")
                        DR = myCLIDs.Tables("ShipData").Select(" CLID = '" & LabelID & "'")
                        If DR.Length > 0 Then
                            For j = 0 To DR.Length - 1
                                DR(j).Delete()
                            Next
                            myCLIDs.Tables("ShipData").AcceptChanges()
                        End If
                    Next
                    myCLIDs.AcceptChanges()
                End If

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("Receiving-ReadShipmentData", LoginData.User.ToUpper, "OSPDJ: " & OrderNo & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function ReadShipCLIDs(ByVal LoginData As ERPLogin, ByVal ShipmentNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs As DataSet = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select CLID from T_CLMaster with (nolock) where ReferenceCLID IS NULL and OrgCode = '{0}' and PurOrdNo = '{1}'", LoginData.OrgCode, ShipmentNo)
                myCLIDs = da.ExecuteDataSet(Sqlstr, "CLIDs")

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("Receiving-ReadShipCLIDs", LoginData.User.ToUpper, "ShipmentNo: " & ShipmentNo & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function GetIRDataFromServer(ByVal myIRData As IRData) As DataSet
        Dim eTraceRS As eTraceWS.eTraceOracleERP = New eTraceWS.eTraceOracleERP
        Dim IRDataRS As eTraceWS.IRData = New eTraceWS.IRData

        Dim SourceServer As String = ""
        Dim myCLIDs As DataSet = Nothing

        Try
            SourceServer = GetOrgServer(myIRData.SourceOrg)
            If SourceServer = "" Then
                Return Nothing
                Exit Function
            End If


            Dim a As Integer
            Dim myUrl As String = eTraceRS.Url

            'a = InStr(8, myUrl, ".", CompareMethod.Text)                     
            a = InStr(8, myUrl, "/", CompareMethod.Text)                      'Find eTrace Server Name with string "/" instead of "."
            If a > 0 Then
                myUrl = Mid(myUrl, 1, 7) & SourceServer & Mid(myUrl, a)
            End If
            eTraceRS.Url = myUrl
            eTraceRS.Timeout = 30 * 60 * 1000

            IRDataRS.User = myIRData.User
            IRDataRS.SourceOrg = myIRData.SourceOrg
            IRDataRS.DestOrg = myIRData.DestOrg
            IRDataRS.ShipmentNo = myIRData.ShipmentNo

            Dim k As Integer = 0

            While (k < 3 And myCLIDs Is Nothing)
                Try
                    myCLIDs = New DataSet
                    myCLIDs = eTraceRS.ReadIRData(IRDataRS)
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("Receiving-GetIRDataFromServer" & Str(k), myIRData.User.ToUpper & " " & Str(k), "ShipmentNo: " & myIRData.ShipmentNo & " with OrgCode " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
                    myCLIDs = Nothing
                End Try
            End While

            Return myCLIDs

        Catch ex As Exception
            ErrorLogging("Receiving-GetIRDataFromServer", myIRData.User.ToUpper, "ShipmentNo: " & myIRData.ShipmentNo & " with OrgCode " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Public Function UpdateRTNoToServer(ByVal myIRData As IRData, ByVal Items As DataSet) As Boolean
        Dim eTraceRS As eTraceWS.eTraceOracleERP = New eTraceWS.eTraceOracleERP
        Dim IRDataRS As eTraceWS.IRData = New eTraceWS.IRData

        UpdateRTNoToServer = False
        Dim SourceServer As String = ""

        Try
            SourceServer = GetOrgServer(myIRData.SourceOrg)
            If SourceServer = "" Then
                Return False
                Exit Function
            End If

            Dim a As Integer
            Dim myUrl As String = eTraceRS.Url

            'a = InStr(8, myUrl, ".", CompareMethod.Text)
            a = InStr(8, myUrl, "/", CompareMethod.Text)                      'Find eTrace Server Name with string "/" instead of "."
            If a > 0 Then
                myUrl = Mid(myUrl, 1, 7) & SourceServer & Mid(myUrl, a)
            End If
            eTraceRS.Url = myUrl
            eTraceRS.Timeout = 30 * 60 * 1000

            IRDataRS.User = myIRData.User
            IRDataRS.SourceOrg = myIRData.SourceOrg
            IRDataRS.DestOrg = myIRData.DestOrg
            IRDataRS.ShipmentNo = myIRData.ShipmentNo
            IRDataRS.IRRTNo = myIRData.IRRTNo

            Dim k As Integer = 0

            While (k < 3 And UpdateRTNoToServer = False)
                Try
                    UpdateRTNoToServer = eTraceRS.UpdateIRRTNo(IRDataRS, Items)
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("Receiving-UpdateRTNoToServer" & Str(k), myIRData.User.ToUpper & " " & Str(k), "IRRTNo: " & myIRData.IRRTNo & " with OrgCode " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
                    UpdateRTNoToServer = False
                End Try
            End While


        Catch ex As Exception
            ErrorLogging("Receiving-UpdateRTNoToServer", myIRData.User.ToUpper, "IRRTNo: " & myIRData.IRRTNo & " with OrgCode " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
            Return False
        End Try

    End Function

    Public Function ReadIRData(ByVal myIRData As IRData) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim myCLIDs As DataSet = New DataSet

            Try
                Dim Sqlstr As String
                Sqlstr = String.Format("Select CLID,MaterialNo,QtyBaseUOM,BaseUOM,DateCode,LotNo,CountryOfOrigin,RTLot,ExpDate,ProdDate,RoHS,MatSuffix1,MatSuffix2,MatSuffix3,Manufacturer,ManufacturerPN,QMLStatus,AddlData,Stemp,MSL,Status=0 from T_CLMaster with (nolock) where (IRRTNo='' or IRRTNo IS NULL) and OrgCode = '{0}' and ShipmentNo = '{1}'", myIRData.SourceOrg, myIRData.ShipmentNo)
                myCLIDs = da.ExecuteDataSet(Sqlstr, "LabelIDs")

                Return myCLIDs

            Catch ex As Exception
                ErrorLogging("Receiving-ReadIRData", myIRData.User.ToUpper, "ShipmentNo: " & myIRData.ShipmentNo & " with DestOrg " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using

    End Function

    Public Function UpdateIRRTNo(ByVal myIRData As IRData, ByVal Items As DataSet) As Boolean
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i As Integer
                Dim ra As Integer
                For i = 0 To Items.Tables(0).Rows.Count - 1
                    Dim LabelID As String = Items.Tables(0).Rows(i)("CLID")
                    ra = da.ExecuteNonQuery(String.Format("update T_CLMaster set ChangedOn=getdate(), ChangedBy='{0}', IRRTNo='{1}' where CLID='{2}'", myIRData.User.ToUpper, myIRData.IRRTNo, LabelID))
                Next

                UpdateIRRTNo = True

            Catch ex As Exception
                ErrorLogging("Receiving-UpdateIRRTNo", myIRData.User.ToUpper, "IRRTNo: " & myIRData.IRRTNo & " with DestOrg " & myIRData.DestOrg & ", " & ex.Message & ex.Source, "E")
                Return False
            End Try

        End Using

    End Function

    Public Function SaveKanbanPO(ByVal Items As DataSet, ByVal BatchID As String, ByVal UserName As String, ByVal StatusCode As Integer) As Boolean
        Using da As DataAccess = GetDataAccess()

            Try
                Dim i As Integer
                Dim sqlstr As String

                i = Items.Tables("po_line").Rows.Count
                If i > 1 Then
                    ErrorLogging("Receiving-Post_KanbanPO", UserName, "BatchID: " & BatchID & " and StatusCode: " & StatusCode & " with KanbanPO Rows; " & i, "I")
                End If

                For i = 0 To Items.Tables("po_line").Rows.Count - 1
                    If Items.Tables("po_line").Rows(i)("po_type") = "KB_PO" Then
                        Dim OrderNo, OrderItem, MaterialNo, KanbanCard, PO_DistributionID, Succ_flag, Return_message, RowState As String
                        Dim RecQty As Decimal

                        If Items.Tables("po_line").Rows(i)("RELEASE_NUM") Is DBNull.Value Then Items.Tables("po_line").Rows(i)("RELEASE_NUM") = 0
                        If Items.Tables("po_line").Rows(i)("SHIPMENT_NUM") Is DBNull.Value Then Items.Tables("po_line").Rows(i)("SHIPMENT_NUM") = 0

                        If Items.Tables("po_line").Rows(i)("RELEASE_NUM") > 0 Then
                            OrderNo = Items.Tables("po_line").Rows(i)("PO_NUMBER").ToString & "-" & Items.Tables("po_line").Rows(i)("RELEASE_NUM").ToString
                        Else
                            OrderNo = Items.Tables("po_line").Rows(i)("PO_NUMBER").ToString
                        End If
                        If Items.Tables("po_line").Rows(i)("SHIPMENT_NUM") > 0 Then
                            OrderItem = Items.Tables("po_line").Rows(i)("LINE_NUM") & "." & Items.Tables("po_line").Rows(i)("SHIPMENT_NUM")
                        Else
                            OrderItem = Items.Tables("po_line").Rows(i)("LINE_NUM")
                        End If
                        OrderItem = OrderItem & "." & Items.Tables("po_line").Rows(i)("DISTRIBUTION_NUM").ToString

                        MaterialNo = Items.Tables("po_line").Rows(i)("ITEM_NUMBER").ToString
                        KanbanCard = Items.Tables("po_line").Rows(i)("KANBAN_CARD_ID").ToString
                        PO_DistributionID = Items.Tables("po_line").Rows(i)("PO_DISTRIBUTION_ID").ToString
                        Succ_flag = Items.Tables("po_line").Rows(i)("succ_flag").ToString
                        Return_message = Items.Tables("po_line").Rows(i)("return_message").ToString

                        RecQty = IIf(Items.Tables("po_line").Rows(i)("REC_PRIMARY_QTY") Is DBNull.Value, 0, Items.Tables(1).Rows(i)("REC_PRIMARY_QTY"))

                        RowState = Items.Tables("po_line").Rows(i).RowState.ToString

                        sqlstr = String.Format("INSERT INTO KanbanPO_tmp (BatchID, PONo, POLine, MaterialNo, RecQty, KanbanCard, PO_DistributionID, Succ_flag, Return_message, UserName, StatusCode, RowState, DateTime ) values ('{0}','{1}','{2}','{3}', '{4}','{5}','{6}','{7}', '{8}', '{9}', '{10}', '{11}', getdate() )", BatchID, OrderNo, OrderItem, MaterialNo, RecQty, KanbanCard, PO_DistributionID, Succ_flag, Return_message, UserName, StatusCode, RowState)
                        da.ExecuteNonQuery(sqlstr)
                    End If
                Next
                SaveKanbanPO = True

            Catch ex As Exception
                ErrorLogging("Receiving-SaveKanbanPO", UserName, ex.Message & ex.Source, "E")
                SaveKanbanPO = False
            End Try

        End Using

    End Function

    Public Function CleanBoxID(ByVal LoginData As ERPLogin, ByVal CLIDs As DataSet) As Boolean
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim BoxID As String = CLIDs.Tables(0).Rows(0)("BoxID")

            Try
                For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                    Dim CLID As String = CLIDs.Tables(0).Rows(i)("CLID")
                    da.ExecuteNonQuery(String.Format("update T_CLMaster set BoxID='', ChangedOn=getDate(), ChangedBy='{0}', LastTransaction='Receiving-CleanBoxID' where CLID='{1}'", LoginData.User, CLID))
                Next
                Return True
            Catch ex As Exception
                ErrorLogging("Receiving-CleanBoxID", LoginData.User, "Clean error for BoxID " & BoxID & ", " & ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
