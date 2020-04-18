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

Public Structure InputData
    Public OrgCode As String
    Public DmdCutDate_Shortage As String          'Demand cut-off date for Shortage calculation
    Public POCutDate_Shortage As String           'PO cut-off date for Shortage calculation
    Public DmdCutDate_Excess As String            'Demand cut-off date for Excess calculation
    Public POCutDate_Excess As String             'PO cut-off date for Excess calculation
    Public SafetyStock As Boolean                 'Include SafetyStock to Shortage or not 
    Public ShortageItems As String                'Input Shortage Item Lists
End Structure


Public Class InventoryOptimization
    Inherits PublicFunction


    Public Function GetMRPData(ByVal myInputData As InputData) As DataSet

        Dim dsItems As DataSet = New DataSet

        Dim ItemData As DataTable
        Dim myDataColumn As DataColumn

        ItemData = New Data.DataTable("ItemData")
        myDataColumn = New Data.DataColumn("OrgCode", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ShortagePN", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ShPNDesc", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ShPNCost", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ShPNBuyer", System.Type.GetType("System.String"))        'Buyer Name
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MFR", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("MPN", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("AMLStatus", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExcessPN", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExPNDesc", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExPNCost", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExPNBuyer", System.Type.GetType("System.String"))        'Buyer Name
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExPNMPQ", System.Type.GetType("System.Double"))          'MPQ
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("SubInventory", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.DateTime"))
        ItemData.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        'ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("OnhandQty", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("ExcessQty", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("AllocatedQty", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("TransferQty", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("InTransit", System.Type.GetType("System.Double"))
        ItemData.Columns.Add(myDataColumn)

        dsItems.Tables.Add(ItemData)


        Dim ErrorTable As DataTable
        Dim ErrorRow As Data.DataRow

        ErrorTable = New Data.DataTable("ErrorTable")
        myDataColumn = New Data.DataColumn("ErrorMsg", System.Type.GetType("System.String"))
        ErrorTable.Columns.Add(myDataColumn)

        dsItems.Tables.Add(ErrorTable)

        Try

            Dim dsMRP As DataSet = New DataSet
            Try
                dsMRP = Oracle_MRPData(myInputData)
            Catch ex As Exception
                ErrorLogging("InventoryOptimization-GetMRPData", "", ex.Message & ex.Source, "E")
                dsMRP = Nothing
            End Try

            'Save or Read MRP data from eTrace is only for test purpose only, will delete later
            ' If dsMRP Is Nothing OrElse dsMRP.Tables.Count = 0 OrElse dsMRP.Tables(0).Rows.Count = 0 Then
            'dsMRP = New DataSet
            'dsMRP = ReadMRPData()         'Read MRP Data from eTrace database for test only
            'ElseIf dsMRP.Tables(0).Rows.Count > 0 Then
            '  SaveMRPData(dsMRP)             'Save MRP Data in eTrace database for test only
            ' End If


            If dsMRP Is Nothing OrElse dsMRP.Tables.Count = 0 OrElse dsMRP.Tables(0).Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No MRP data found from Oracle"
                ErrorTable.Rows.Add(ErrorRow)
                Return dsItems
                Exit Function
            End If


            Dim i, j, k, n As Integer
            Dim dsExcess As DataSet = New DataSet
            Dim dsShortage As DataSet = New DataSet

            Dim MatlList As DataTable
            Dim myDataRow As Data.DataRow

            MatlList = New Data.DataTable("MatlList")
            myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Description", System.Type.GetType("System.String"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("StdCost", System.Type.GetType("System.Double"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Buyer", System.Type.GetType("System.String"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("MPQ", System.Type.GetType("System.Double"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Quantity", System.Type.GetType("System.Double"))
            MatlList.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("SfyStock", System.Type.GetType("System.Double"))
            MatlList.Columns.Add(myDataColumn)

            dsExcess.Tables.Add(MatlList)
            dsShortage = dsExcess.Copy

            Dim DR() As DataRow
            Dim ItemList() As String
            Dim MaterialNo As String

            j = dsMRP.Tables(0).Rows.Count
            ReDim ItemList(j)
            ItemList.Initialize()
            k = 0

            'For i = 0 To dsMRP.Tables(0).Rows.Count - 1
            '    If dsMRP.Tables(0).Rows(i)("Quantity") IsNot DBNull.Value And Val(dsMRP.Tables(0).Rows(i)("Quantity")) > 0 Then          'Get Excess Items 
            '        MaterialNo = dsMRP.Tables(0).Rows(i)("Item_name")
            '        DR = Nothing
            '        DR = dsExcess.Tables("MatlList").Select(" MaterialNo = '" & MaterialNo & "'")
            '        If DR.Length = 0 Then
            '            myDataRow = dsExcess.Tables("MatlList").NewRow()
            '            myDataRow("MaterialNo") = MaterialNo
            '            myDataRow("Description") = dsMRP.Tables(0).Rows(i)("Description").ToString
            '            myDataRow("Buyer") = dsMRP.Tables(0).Rows(i)("Buyer").ToString
            '            myDataRow("StdCost") = IIf(dsMRP.Tables(0).Rows(i)("Standard_cost") Is DBNull.Value, 0, dsMRP.Tables(0).Rows(i)("Standard_cost"))
            '            myDataRow("SfyStock") = IIf(dsMRP.Tables(0).Rows(i)("Safety_qty") Is DBNull.Value, 0, dsMRP.Tables(0).Rows(i)("Safety_qty"))
            '            myDataRow("Quantity") = Val(dsMRP.Tables(0).Rows(i)("Quantity"))
            '            dsExcess.Tables("MatlList").Rows.Add(myDataRow)

            '            ItemList(k) = MaterialNo                              'Prepare Excess Items for AML Lists reading
            '            k = k + 1
            '        End If
            '    End If
            'Next

            Dim Qty As Double = 0
            Dim ShortageItems, sqlstr As String
            ShortageItems = myInputData.ShortageItems.ToString

            sqlstr = " Quantity < '" & Qty & "'"                              'Find Shortage Items
            If ShortageItems = "" Or ShortageItems = "*" Then
            Else                                                              'Check if the ShortagePN exists in the Input Item Lists or not
                If ShortageItems.Contains("*") Then
                    sqlstr = sqlstr & " and " & "Item_name like '" & ShortageItems.Replace("*", "%") & "' "
                Else
                    sqlstr = sqlstr & " and " & "Item_name in ( '" & ShortageItems.Replace(",", "','") & "') "
                End If
            End If

            DR = Nothing
            DR = dsMRP.Tables(0).Select(sqlstr)
            If DR.Length > 0 Then
                For i = 0 To DR.Length - 1
                    Dim Shortage, SafetyStock As Double
                    Shortage = IIf(DR(i)("Quantity") Is DBNull.Value, 0, DR(i)("Quantity"))
                    SafetyStock = IIf(DR(i)("Safety_qty") Is DBNull.Value, 0, DR(i)("Safety_qty"))

                    If myInputData.SafetyStock = False Then
                        Shortage = Shortage + SafetyStock
                    End If

                    If Shortage < 0 Then
                        myDataRow = dsShortage.Tables("MatlList").NewRow()
                        myDataRow("MaterialNo") = DR(i)("Item_name")
                        myDataRow("Description") = DR(i)("Description").ToString
                        myDataRow("StdCost") = IIf(DR(i)("Standard_cost") Is DBNull.Value, 0, DR(i)("Standard_cost"))
                        myDataRow("Buyer") = DR(i)("Buyer").ToString
                        myDataRow("Quantity") = 0 - Shortage
                        myDataRow("SfyStock") = SafetyStock
                        dsShortage.Tables("MatlList").Rows.Add(myDataRow)

                        ItemList(k) = DR(i)("Item_name").ToString              'Prepare Shortage Items for AML Lists reading
                        k = k + 1
                    End If
                Next
            End If

            If dsShortage Is Nothing OrElse dsShortage.Tables.Count = 0 OrElse dsShortage.Tables(0).Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No available Shortage Items found, pls check your input"
                ErrorTable.Rows.Add(ErrorRow)
                Return dsItems
                Exit Function
            End If

            Dim dsAML As DataSet = New DataSet
            Try
                dsAML = GetAML(ItemList)                    'from iPro
            Catch ex As Exception
                ErrorLogging("InventoryOptimization-GetAML", "", ex.Message & ex.Source, "E")
                dsAML = Nothing
            End Try

            If dsAML Is Nothing OrElse dsAML.Tables.Count = 0 OrElse dsAML.Tables(0).Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No AML data found from iPro"
                ErrorTable.Rows.Add(ErrorRow)
                Return dsItems
                Exit Function
            End If


            Dim dsMPN As DataSet = New DataSet
            Try
                dsMPN = GeteTraceData()
            Catch ex As Exception
                ErrorLogging("InventoryOptimization-GeteTraceData", "", ex.Message & ex.Source, "E")
                dsMPN = Nothing
            End Try

            If dsMPN Is Nothing OrElse dsMPN.Tables.Count = 0 OrElse dsMPN.Tables(0).Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No ManufacturerPN data found in eTrace"
                ErrorTable.Rows.Add(ErrorRow)
                Return dsItems
                Exit Function
            End If

            Dim CLIDdr(), AMLdr(), MRPdr() As DataRow
            Dim ShortagePN, ExcessPN, RTLot, MFR, MPN, AMLstatus As String


            For i = 0 To dsShortage.Tables(0).Rows.Count - 1
                ShortagePN = dsShortage.Tables(0).Rows(i)("MaterialNo").ToString

                AMLdr = Nothing
                AMLdr = dsAML.Tables("AMLData").Select(" MaterialNo = '" & ShortagePN & "'")
                If AMLdr.Length > 0 Then
                    For j = 0 To AMLdr.Length - 1
                        MFR = AMLdr(j)("MFR").ToString
                        MPN = AMLdr(j)("MPN").ToString
                        AMLstatus = AMLdr(j)("AMLStatus").ToString

                        CLIDdr = Nothing
                        CLIDdr = dsMPN.Tables(0).Select(" Manufacturer = '" & MFR & "' and ManufacturerPN = '" & MPN & "'")
                        If CLIDdr.Length > 0 Then
                            For k = 0 To CLIDdr.Length - 1

                                ExcessPN = CLIDdr(k)("MaterialNo").ToString
                                RTLot = CLIDdr(k)("RTLot").ToString
                                Dim ExPNMPQ As Double = Val(CLIDdr(k)("MPQ"))

                                ' MRPdr = Nothing
                                'MRPdr = dsMRP.Tables("mrpdata").Select(" Item_name = '" & ExcessPN & "' and Lot_number = '" & RTLot & "'")
                                'If MRPdr.Length > 0 Then
                                'Dim ExPNDesc As String = MRPdr(0)("Description").ToString
                                'Dim ExPNBuyer As String = MRPdr(0)("Buyer").ToString
                                'Dim ExPNCost As Double = IIf(MRPdr(0)("Standard_cost") Is DBNull.Value, 0, MRPdr(0)("Standard_cost"))
                                'Dim ExcessQty As Double = IIf(MRPdr(0)("Quantity") Is DBNull.Value, 0, MRPdr(0)("Quantity"))


                                Dim ExPNDesc As String = CLIDdr(k)("Description").ToString
                                Dim ExPNBuyer As String = CLIDdr(k)("Buyer").ToString
                                Dim ExPNCost As Double = IIf(CLIDdr(k)("Standard_cost") Is DBNull.Value, 0, CLIDdr(k)("Standard_cost"))
                                Dim ExcessQty As Double = IIf(CLIDdr(k)("Quantity") Is DBNull.Value, 0, CLIDdr(k)("Quantity"))
                                Dim TransferQty As Double = IIf(CLIDdr(k)("TransferQty") Is DBNull.Value, 0, CLIDdr(k)("TransferQty"))

                                If (TransferQty > ExcessQty) Then
                                    TransferQty = ExcessQty
                                End If

                                'For n = 0 To MRPdr.Length - 1



                                myDataRow = dsItems.Tables("ItemData").NewRow()
                                myDataRow("OrgCode") = myInputData.OrgCode
                                myDataRow("ShortagePN") = ShortagePN
                                myDataRow("ShPNDesc") = dsShortage.Tables(0).Rows(i)("Description").ToString
                                myDataRow("ShPNCost") = IIf(dsShortage.Tables(0).Rows(i)("StdCost") Is DBNull.Value, 0, dsShortage.Tables(0).Rows(i)("StdCost"))
                                myDataRow("ShPNBuyer") = dsShortage.Tables(0).Rows(i)("Buyer").ToString
                                myDataRow("Shortage") = Val(dsShortage.Tables(0).Rows(i)("Quantity"))
                                myDataRow("MFR") = MFR
                                myDataRow("MPN") = MPN
                                myDataRow("AMLstatus") = AMLstatus
                                myDataRow("ExcessPN") = ExcessPN
                                myDataRow("ExPNDesc") = ExPNDesc
                                myDataRow("ExPNCost") = ExPNCost
                                myDataRow("ExPNBuyer") = ExPNBuyer
                                myDataRow("ExPNMPQ") = ExPNMPQ

                                myDataRow("SubInventory") = CLIDdr(k)("SubInventory").ToString
                                myDataRow("Locator") = CLIDdr(k)("Locator").ToString
                                'myDataRow("RTLot") = RTLot

                                If Not CLIDdr(k)("expiration_date") Is DBNull.Value Then
                                    myDataRow("ExpDate") = CDate(CLIDdr(k)("expiration_date"))
                                End If

                                myDataRow("ExcessQty") = ExcessQty
                                myDataRow("OnhandQty") = IIf(CLIDdr(k)("Onhand_qty") Is DBNull.Value, 0, CLIDdr(k)("Onhand_qty"))
                                myDataRow("AllocatedQty") = IIf(CLIDdr(k)("Allocated_qty") Is DBNull.Value, 0, CLIDdr(k)("Allocated_qty"))
                                myDataRow("AvlQty") = IIf(CLIDdr(k)("Ava_qty") Is DBNull.Value, 0, CLIDdr(k)("Ava_qty"))
                                myDataRow("TransferQty") = TransferQty
                                myDataRow("InTransit") = 0

                                dsItems.Tables("ItemData").Rows.Add(myDataRow)



                                If CLIDdr(k)("In_transit") IsNot DBNull.Value And Val(CLIDdr(k)("In_transit")) > 0 Then
                                    myDataRow = dsItems.Tables("ItemData").NewRow()
                                    myDataRow("OrgCode") = myInputData.OrgCode
                                    myDataRow("ShortagePN") = ShortagePN
                                    myDataRow("ShPNDesc") = dsShortage.Tables(0).Rows(i)("Description").ToString
                                    myDataRow("ShPNCost") = IIf(dsShortage.Tables(0).Rows(i)("StdCost") Is DBNull.Value, 0, dsShortage.Tables(0).Rows(i)("StdCost"))
                                    myDataRow("ShPNBuyer") = dsShortage.Tables(0).Rows(i)("Buyer").ToString
                                    myDataRow("Shortage") = Val(dsShortage.Tables(0).Rows(i)("Quantity"))
                                    myDataRow("MFR") = MFR
                                    myDataRow("MPN") = MPN
                                    myDataRow("AMLstatus") = AMLstatus
                                    myDataRow("ExcessPN") = ExcessPN
                                    myDataRow("ExPNDesc") = ExPNDesc
                                    myDataRow("ExPNCost") = ExPNCost
                                    myDataRow("ExPNBuyer") = ExPNBuyer
                                    myDataRow("ExPNMPQ") = ExPNMPQ
                                    myDataRow("SubInventory") = ""
                                    myDataRow("Locator") = ""
                                    'myDataRow("RTLot") = ""

                                    'If Not CLIDdr(k)("ExpDate") Is DBNull.Value Then
                                    '    myDataRow("ExpDate") = CDate(CLIDdr(k)("ExpDate"))
                                    'End If

                                    myDataRow("ExcessQty") = ExcessQty
                                    myDataRow("OnhandQty") = 0
                                    myDataRow("AllocatedQty") = 0
                                    myDataRow("AvlQty") = 0
                                    myDataRow("TransferQty") = 0
                                    myDataRow("InTransit") = Val(CLIDdr(k)("In_transit"))

                                    dsItems.Tables("ItemData").Rows.Add(myDataRow)
                                End If


                                'Next

                                'Add InTransit Qty from Oracle if there exists

                                'End If

                            Next
                        End If
                    Next
                End If
            Next

            ' Record error message if no data found 
            If dsItems.Tables("ItemData").Rows.Count = 0 Then
                ErrorRow = ErrorTable.NewRow()
                ErrorRow("ErrorMsg") = "No ManufacturerPN data found in eTrace"
                ErrorTable.Rows.Add(ErrorRow)
                Return dsItems
                Exit Function
            End If

            Return dsItems

        Catch ex As Exception
            ErrorLogging("InventoryOptimization-GetMRPData", "", ex.Message & ex.Source, "E")
            Return Nothing
        End Try

    End Function

    Public Function GeteTraceData() As DataSet

        Using da As DataAccess = GetDataAccess()
            Dim sql1 As String
            Try
                Dim ds As New DataSet()
                ds.Tables.Add("ItemList")
                da.ExecuteNonQuery(String.Format("exec dbo.sp_get_etr_data"))

                ds = da.ExecuteDataSet(String.Format("select * from dbo.g_etr_data"))
                Return ds
            Catch oe As Exception
                ErrorLogging("InventoryOptimization-Oracle_MRPData", "", oe.Message & oe.Source, "E")
                Return Nothing
            End Try

        End Using



        ' Using da As DataAccess = GetDataAccess()
        'Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTracePDConnString"))
        'Dim myCommand As SqlClient.SqlCommand
        'Dim dsMPN = New DataSet
        'Dim ItemList As DataTable

        'Dim myDataRow As Data.DataRow
        'Dim myDataColumn As DataColumn

        'ItemList = New Data.DataTable("ItemList")
        'myDataColumn = New Data.DataColumn("MaterialNo", System.Type.GetType("System.String"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.DateTime"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.Double"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("InTransit", System.Type.GetType("System.Double"))
        'ItemList.Columns.Add(myDataColumn)
        'myDataColumn = New Data.DataColumn("MPQ", System.Type.GetType("System.Double"))
        'ItemList.Columns.Add(myDataColumn)


        'dsMPN.Tables.Add(ItemList)

        'Dim dsExcess As DataSet = New DataSet
        'dsExcess = dsMPN.copy

        'Dim ds As DataSet = New DataSet
        'Dim myCLIDs As DataSet = New DataSet

        'Dim i, j As Integer
        'Dim DR() As DataRow
        'Dim MaterialNo, RTLot As String
        'Dim MaxCLIDQty As Double

        'Try
        '    'For i = 0 To Items.Tables(0).Rows.Count - 1
        '    For i = 0 To 100
        '        If Items.Tables(0).Rows(i)("Quantity") IsNot DBNull.Value And Val(Items.Tables(0).Rows(i)("Quantity")) > 0 Then         'Read CLID info with Excess PN
        '            MaterialNo = Items.Tables(0).Rows(i)("Item_name").ToString
        '            RTLot = Items.Tables(0).Rows(i)("Lot_number").ToString

        '            MaxCLIDQty = 0
        '            If MaterialNo <> "" And RTLot <> "" Then
        '                DR = Nothing
        '                DR = dsExcess.Tables("ItemList").Select(" MaterialNo = '" & MaterialNo & "' and RTLot = '" & RTLot & "'")
        '                If DR.Length = 0 Then
        '                    myDataRow = dsExcess.Tables("ItemList").NewRow()
        '                    myDataRow("MaterialNo") = MaterialNo
        '                    myDataRow("RTLot") = RTLot

        '                    ds = New DataSet
        '                    Dim Sqlstr As String
        '                    Sqlstr = String.Format("Select CLID,MaterialNo,SLOC,StorageBin,RTLot,ExpDate,Manufacturer,ManufacturerPN,QtyBaseUOM,BaseUOM,StatusCode from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and RTLot = '{2}' ", myInputData.OrgCode, MaterialNo, RTLot)


        '                    Dim _SCmd As New SqlCommand(Sqlstr, myConn)
        '                    Dim _SApt As New SqlDataAdapter(_SCmd)

        '                    _SApt.Fill(ds, "LabelIDs")


        '                    'ds = da.ExecuteDataSet(Sqlstr, "LabelIDs")

        '                    If ds.Tables(0).Rows.Count > 0 Then
        '                        If myCLIDs.Tables.Count = 0 Then
        '                            myCLIDs = ds.Clone
        '                        End If

        '                        For Each drw As DataRow In ds.Tables("LabelIDs").Rows
        '                            myCLIDs.Tables("LabelIDs").ImportRow(drw)

        '                            'Find the Max CLID Qty as Item MPQ
        '                            If MaxCLIDQty < Val(drw("QtyBaseUOM")) Then
        '                                MaxCLIDQty = Val(drw("QtyBaseUOM"))
        '                            End If
        '                        Next
        '                    End If

        '                    myDataRow("MPQ") = MaxCLIDQty
        '                    dsExcess.Tables("ItemList").Rows.Add(myDataRow)

        '                End If
        '            End If
        '        End If
        '    Next
        '    myConn.Close()
        'Catch ex As Exception
        '    myConn.Close()
        '    ErrorLogging("InventoryOptimization-GetCLIDList", "", ex.Message & ex.Source, "E")
        '    Return Nothing
        'End Try


        'Try
        '    For i = 0 To myCLIDs.Tables(0).Rows.Count - 1
        '        Dim MFR, MPN As String
        '        MaterialNo = myCLIDs.Tables(0).Rows(i)("MaterialNo").ToString
        '        RTLot = myCLIDs.Tables(0).Rows(i)("RTLot").ToString
        '        MFR = myCLIDs.Tables(0).Rows(i)("Manufacturer").ToString
        '        MPN = myCLIDs.Tables(0).Rows(i)("ManufacturerPN").ToString

        '        MaxCLIDQty = 0
        '        DR = Nothing
        '        DR = dsExcess.Tables("ItemList").Select(" MaterialNo = '" & MaterialNo & "'")
        '        If DR.Length > 0 Then
        '            For j = 0 To DR.Length - 1
        '                If MaxCLIDQty < Val(DR(j)("MPQ")) Then
        '                    MaxCLIDQty = Val(DR(j)("MPQ"))
        '                End If
        '            Next
        '        End If

        '        If MFR <> "" And MPN <> "" Then
        '            DR = Nothing
        '            DR = dsMPN.Tables(0).Select(" MaterialNo = '" & MaterialNo & "' and RTLot = '" & RTLot & "' and Manufacturer = '" & MFR & "' and ManufacturerPN = '" & MPN & "'")
        '            If DR.Length = 0 Then
        '                myDataRow = dsMPN.Tables("ItemList").NewRow()
        '                myDataRow("MaterialNo") = MaterialNo
        '                myDataRow("RTLot") = RTLot

        '                If Not myCLIDs.Tables(0).Rows(i)("ExpDate") Is DBNull.Value Then
        '                    myDataRow("ExpDate") = CDate(myCLIDs.Tables(0).Rows(i)("ExpDate"))
        '                End If

        '                myDataRow("Manufacturer") = MFR
        '                myDataRow("ManufacturerPN") = MPN
        '                myDataRow("AvlQty") = 0
        '                myDataRow("InTransit") = 0
        '                myDataRow("MPQ") = MaxCLIDQty

        '                If myCLIDs.Tables(0).Rows(i)("StatusCode") = 1 Then
        '                    If myCLIDs.Tables(0).Rows(i)("SLOC").ToString = "" Then
        '                        myDataRow("InTransit") = Val(myCLIDs.Tables(0).Rows(i)("QtyBaseUOM"))
        '                    Else
        '                        myDataRow("AvlQty") = Val(myCLIDs.Tables(0).Rows(i)("QtyBaseUOM"))
        '                    End If
        '                End If
        '                dsMPN.Tables("ItemList").Rows.Add(myDataRow)

        '            Else
        '                Dim AvlQty, InTransit As Double
        '                AvlQty = Val(DR(0)("AvlQty"))
        '                InTransit = Val(DR(0)("InTransit"))

        '                If myCLIDs.Tables(0).Rows(i)("StatusCode") = 1 Then
        '                    If myCLIDs.Tables(0).Rows(i)("SLOC").ToString = "" Then
        '                        DR(0)("InTransit") = Val(myCLIDs.Tables(0).Rows(i)("QtyBaseUOM")) + InTransit
        '                    Else
        '                        DR(0)("AvlQty") = Val(myCLIDs.Tables(0).Rows(i)("QtyBaseUOM")) + AvlQty
        '                    End If
        '                End If
        '                DR(0).AcceptChanges()
        '            End If
        '            dsMPN.tables(0).AcceptChanges()

        '        End If
        '    Next

        '    Return dsMPN

        'Catch ex As Exception
        '    ErrorLogging("InventoryOptimization-GeteTraceData", "", ex.Message & ex.Source, "E")
        '    Return Nothing
        'End Try

        '  End Using


    End Function

    Public Function Oracle_MRPData(ByVal myInputData As InputData) As DataSet
        Using da As DataAccess = GetDataAccess()

            ' Dim oda As OracleDataAdapter = da.Ascp_Sele()
            'Dim thisConnection As New SqlConnection(ConfigurationManager.ConnectionStrings("eTracePDConnectionString").ConnectionString)
            Dim sql1 As String
            'Dim PlanName As String = "FROZENMRP"         'MRP Plan Name for Raw Components: FROZENMRP, requested by Huaiyuan Xu
            'Dim PlanName As String = "MRP-SAT"            'MRP Plan Name for Raw Components: MRP-SAT
            ' Dim CategorySetID As Integer = 4              'EMR_INV:  Inventory Classification Category Set

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("mrpdata")

                'exec(dbo.sp_Get_mrp_data) '405','2010/10/30','2010/12/31',null,null 

                sql1 = String.Format("exec dbo.sp_Get_mrp_data '{0}','{1}','{2}','{3}','{4}'", myInputData.OrgCode, myInputData.DmdCutDate_Shortage, myInputData.DmdCutDate_Excess, myInputData.POCutDate_Shortage, myInputData.POCutDate_Excess)

                'Dim thisAdapter1 As New SqlDataAdapter(sql1, thisConnection)
                'thisAdapter1.SelectCommand.CommandTimeout = 60 * 60 * 30
                'thisAdapter1.SelectCommand.ExecuteNonQuery()

                da.ExecuteNonQuery(sql1)


                'oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'oda.SelectCommand.CommandText = "apps.xxetr_ascp_pkg.read_ascp_result"
                'oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar, 50).Value = myInputData.OrgCode
                'oda.SelectCommand.Parameters.Add("p_plan_name", OracleType.VarChar, 200).Value = PlanName
                'oda.SelectCommand.Parameters.Add("p_category_set_id", OracleType.Int32).Value = CategorySetID     'EMR_INV
                'oda.SelectCommand.Parameters.Add("p_dmdate_s", OracleType.VarChar, 50).Value = myInputData.DmdCutDate_Shortage
                'oda.SelectCommand.Parameters.Add("p_podate_s", OracleType.VarChar, 50).Value = myInputData.POCutDate_Shortage
                'oda.SelectCommand.Parameters.Add("p_dmdate_e", OracleType.VarChar, 50).Value = myInputData.DmdCutDate_Excess
                'oda.SelectCommand.Parameters.Add("p_podate_e", OracleType.VarChar, 50).Value = myInputData.POCutDate_Excess
                'oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output

                'oda.SelectCommand.Parameters.Add("o_success_flag", OracleType.VarChar, 50)
                'oda.SelectCommand.Parameters.Add("o_error_mssg", OracleType.VarChar, 8000)
                'oda.SelectCommand.Parameters("o_success_flag").Direction = ParameterDirection.Output
                'oda.SelectCommand.Parameters("o_error_mssg").Direction = ParameterDirection.Output
                'oda.SelectCommand.Connection.Open()

                'oda.Fill(ds, "mrpdata")

                'oda.SelectCommand.Connection.Close()
                ds = ReadMRPData()
                Return ds
            Catch oe As Exception
                ErrorLogging("InventoryOptimization-Oracle_MRPData", "", oe.Message & oe.Source, "E")
                Return Nothing
            End Try
        End Using

    End Function

    Public Function SaveMRPData(ByVal Items As DataSet) As Boolean
        Using da As DataAccess = GetDataAccess()

            SaveMRPData = False
            If Items.Tables(0).TableName.ToUpper <> "MRPDATA" OrElse Items.Tables(0).Columns(0).ColumnName.ToUpper <> "ITEM_NAME" Then Exit Function

            Try
                Dim i As Integer
                Dim sqlstr As String
                sqlstr = String.Format("DELETE from MRP_tmp ")
                da.ExecuteNonQuery(sqlstr)

                For i = 0 To Items.Tables(0).Rows.Count - 1
                    Dim Item_name, Revision, Lot_number, SubInventory, Locator, Buyer As String
                    Dim Standard_cost, Safety_qty, Onhand_qty, Ava_qty, Allocated_qty, Quantity, In_Transit As Double

                    Item_name = Items.Tables(0).Rows(i)("Item_name").ToString
                    Revision = Items.Tables(0).Rows(i)("Revision").ToString
                    Lot_number = Items.Tables(0).Rows(i)("Lot_number").ToString
                    SubInventory = Items.Tables(0).Rows(i)("SubInventory").ToString
                    Locator = Items.Tables(0).Rows(i)("Locator").ToString

                    Standard_cost = IIf(Items.Tables(0).Rows(i)("Standard_cost") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Standard_cost"))
                    Safety_qty = IIf(Items.Tables(0).Rows(i)("Safety_qty") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Safety_qty"))

                    Onhand_qty = IIf(Items.Tables(0).Rows(i)("Onhand_qty") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Onhand_qty"))
                    Ava_qty = IIf(Items.Tables(0).Rows(i)("Ava_qty") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Ava_qty"))
                    Allocated_qty = IIf(Items.Tables(0).Rows(i)("Allocated_qty") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Allocated_qty"))
                    Quantity = IIf(Items.Tables(0).Rows(i)("Quantity") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("Quantity"))
                    Buyer = Items.Tables(0).Rows(i)("Buyer").ToString

                    In_Transit = IIf(Items.Tables(0).Rows(i)("In_Transit") Is DBNull.Value, 0, Items.Tables(0).Rows(i)("In_Transit"))

                    sqlstr = String.Format("INSERT INTO MRP_tmp (Item_name, Revision, Standard_cost, Lot_number, SubInventory, Locator, Safety_qty, Onhand_qty, Ava_qty, Allocated_qty, Quantity, Buyer, In_Transit) values ('{0}','{1}','{2}','{3}', '{4}','{5}','{6}','{7}', '{8}','{9}','{10}','{11}','{12}')", Item_name, Revision, Standard_cost, Lot_number, SubInventory, Locator, Safety_qty, Onhand_qty, Ava_qty, Allocated_qty, Quantity, Buyer, In_Transit)
                    da.ExecuteNonQuery(sqlstr)

                Next
                SaveMRPData = True

            Catch ex As Exception
                ErrorLogging("InventoryOptimization-SaveMRPData", "", ex.Message & ex.Source, "E")
                SaveMRPData = False
            End Try

        End Using


    End Function

    Public Function ReadMRPData() As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim ds As DataSet = New DataSet

                Dim sqlstr As String
                sqlstr = String.Format("Select * from MRP_tmp ")
                ds = da.ExecuteDataSet(sqlstr, "mrpdata")

                Return ds

            Catch ex As Exception
                ErrorLogging("InventoryOptimization-ReadMRPData", "", ex.Message & ex.Source, "E")
                Return Nothing
            End Try

        End Using


    End Function

End Class
