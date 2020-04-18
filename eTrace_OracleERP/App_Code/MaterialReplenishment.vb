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

Public Structure ProdPickingStructure
    Public Type As Integer
    Public PO As String
    Public DJQty As String
    Public DJRev As String
    Public BuildQty As String
    Public ProdFloor As String
    Public ProdLine As String
    Public ETA As Date
    'Public ProdStation As String
    Public SupplyType As String
    Public ReasonCode As String
    Public DestSubInv As String
    Public DestLocator As String
    Public CheckSubInv As String
    Public CheckLocator As String
    Public Product As String
    Public Status As String
    Public MakeBuy As String
    Public SrcSubInv As String
End Structure

Public Structure ProdPickingResponse
    Public PDTO As String
    Public ErrMessage As String
End Structure

Public Structure PPDataRst
    Public ErrMsg As String
    Public dsList As DataSet
End Structure

Public Structure StrTemplate
    Public Template As String
    Public Desc As String
    Public LastChangedBy As String
    Public LastChangedOn As String
    Public ErrMesg As String
End Structure

Public Structure Get_Qty
    Public AvlQty As Decimal
    Public WHSQty As Decimal
End Structure

'Public Structure CreateCLIDforPOLabel
'    Public CLID As String
'    Public PCBA As String
'    Public POID As String
'    Public Qty As String
'    Public UOM As String
'    Public RecDate As String
'    Public OrgCode As String
'End Structure

Public Structure OrderInfo
    Public MatlNo As String
    Public OrderQty As Double
    Public OpenQty As Double
    Public ErrorMsg As String
End Structure

Public Structure BOMStr
    Public MatlNo As String
    Public Qty As Double
End Structure

Public Structure ProductLabel
    Public MatlNo As String
    'Public MatlDesc As String
    Public CLID As String
    Public Qty As Double
    'Public StartDate As Date
    Public ErrorMsg As String
End Structure

Public Structure dj_response
    Public flag As String
    Public errormsg As String
    Public dsCLID As DataSet
    'Public subInv As String
    'Public locator As String
End Structure

Public Class MaterialReplemenishment
    Inherits PublicFunction
    Public ReturnTemplate As DataSet = New DataSet

    Public Sub New()
    End Sub

    Public Function SaveProdPicking(ByVal TOMydataSet As DataSet, ByVal TemplateStr As StrTemplate, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As String
        Dim j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer


        Try
            myConn.Open()
            If Header.SupplyType = "Pull" Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateHeader (TOTemplate,Description,Product,ProdFloor,LastChangedBy,LastChangedOn,SupplyType,DestSubInv,ReasonCode) values (@TOTemplate,@Description,@Product,@ProdFloor,@LastChangedBy,getdate(),@SupplyType,@DestSubInv,@ReasonCode)", myConn)
                CLMasterSQLCommand.Parameters.Add("@DestSubInv", SqlDbType.VarChar, 50, "DestSubInv")
                CLMasterSQLCommand.Parameters("@DestSubInv").Value = GetOrderInfoFromERP(ERPLoginData, TOMydataSet.Tables("POItems").Rows(0)("PONo")).Tables(0).Rows(0)(5)
            Else
                CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateHeader (TOTemplate,Description,Product,ProdFloor,LastChangedBy,LastChangedOn,SupplyType,ReasonCode) values (@TOTemplate,@Description,@Product,@ProdFloor,@LastChangedBy,getdate(),@SupplyType,@ReasonCode)", myConn)
            End If
            CLMasterSQLCommand.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand.Parameters.Add("@Description", SqlDbType.NVarChar, 100, "Description")
            CLMasterSQLCommand.Parameters.Add("@Product", SqlDbType.VarChar, 50, "Product")
            CLMasterSQLCommand.Parameters.Add("@ProdFloor", SqlDbType.VarChar, 50, "ProdFloor")
            'CLMasterSQLCommand.Parameters.Add("@ProdStation", SqlDbType.VarChar, 50, "ProdStation")
            CLMasterSQLCommand.Parameters.Add("@LastChangedBy", SqlDbType.VarChar, 50, "LastChangedBy")
            CLMasterSQLCommand.Parameters.Add("@SupplyType", SqlDbType.VarChar, 20, "SupplyType")
            CLMasterSQLCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")

            CLMasterSQLCommand1 = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateItem (TOTemplate,Item,Material,Manufacturer,MPN,Qty,MaterialRevision) values (@TOTemplate,@Item,@Material,@Manufacturer,@MPN,@Qty,@MaterialRevision)", myConn)

            CLMasterSQLCommand1.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand1.Parameters.Add("@Item", SqlDbType.Int, 4, "Item")
            CLMasterSQLCommand1.Parameters.Add("@Material", SqlDbType.VarChar, 50, "Material")
            CLMasterSQLCommand1.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 100, "Manufacturer")
            CLMasterSQLCommand1.Parameters.Add("@MPN", SqlDbType.VarChar, 100, "MPN")
            CLMasterSQLCommand1.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
            CLMasterSQLCommand1.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 50, "MaterialRevision")

            CLMasterSQLCommand.Parameters("@TOTemplate").Value = TemplateStr.Template
            CLMasterSQLCommand.Parameters("@Description").Value = TemplateStr.Desc
            CLMasterSQLCommand.Parameters("@ProdFloor").Value = Header.ProdFloor
            'CLMasterSQLCommand.Parameters("@ProdStation").Value = Header.ProdStation
            CLMasterSQLCommand.Parameters("@LastChangedBy").Value = ERPLoginData.User.ToUpper
            CLMasterSQLCommand.Parameters("@SupplyType").Value = Header.SupplyType
            CLMasterSQLCommand.Parameters("@ReasonCode").Value = Header.ReasonCode

            CLMasterSQLCommand.Parameters("@Product").Value = TOMydataSet.Tables("POItems").Rows(0)("Style")
            CLMasterSQLCommand.CommandType = CommandType.Text
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            For j = 0 To TOMydataSet.Tables("ItemDetails").Rows.Count - 1
                CLMasterSQLCommand1.Parameters("@TOTemplate").Value = TemplateStr.Template
                CLMasterSQLCommand1.Parameters("@Item").Value = j + 1
                CLMasterSQLCommand1.Parameters("@Material").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Mat")
                CLMasterSQLCommand1.Parameters("@MPN").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("ManufacturerPN")
                CLMasterSQLCommand1.Parameters("@Manufacturer").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Manufacturer")
                CLMasterSQLCommand1.Parameters("@Qty").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("PerQty")
                CLMasterSQLCommand1.Parameters("@MaterialRevision").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Revision")
                CLMasterSQLCommand1.CommandType = CommandType.Text
                CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand1.ExecuteNonQuery()
            Next
            myConn.Close()

            SaveProdPicking = "okay"
        Catch ex As Exception
            ErrorLogging("MMC-SaveProdPicking", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            SaveProdPicking = "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Save_Template(ByVal TOMydataSet As DataSet, ByVal TemplateStr As StrTemplate, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As String
        Dim j As Integer
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer


        Try
            myConn.Open()
            If Header.SupplyType <> "Push" Then
                CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateHeader (TOTemplate,Description,Product,ProdFloor,LastChangedBy,LastChangedOn,SupplyType,DestSubInv,DestLocator,ReasonCode) values (@TOTemplate,@Description,@Product,@ProdFloor,@LastChangedBy,getdate(),@SupplyType,@DestSubInv,@DestLocator,@ReasonCode)", myConn)
                CLMasterSQLCommand.Parameters.Add("@DestSubInv", SqlDbType.VarChar, 50, "DestSubInv")
                'CLMasterSQLCommand.Parameters("@DestSubInv").Value = GetOrderInfoFromERP(ERPLoginData, TOMydataSet.Tables("POItems").Rows(0)("DJNo")).Tables(0).Rows(0)(5)
                CLMasterSQLCommand.Parameters("@DestSubInv").Value = Header.CheckSubInv
                CLMasterSQLCommand.Parameters.Add("@DestLocator", SqlDbType.VarChar, 50, "DestLocator")
                CLMasterSQLCommand.Parameters("@DestLocator").Value = Header.CheckLocator
            Else
                CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateHeader (TOTemplate,Description,Product,ProdFloor,LastChangedBy,LastChangedOn,SupplyType,ReasonCode) values (@TOTemplate,@Description,@Product,@ProdFloor,@LastChangedBy,getdate(),@SupplyType,@ReasonCode)", myConn)
            End If
            CLMasterSQLCommand.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand.Parameters.Add("@Description", SqlDbType.NVarChar, 100, "Description")
            CLMasterSQLCommand.Parameters.Add("@Product", SqlDbType.VarChar, 50, "Product")
            CLMasterSQLCommand.Parameters.Add("@ProdFloor", SqlDbType.VarChar, 50, "ProdFloor")
            'CLMasterSQLCommand.Parameters.Add("@ProdStation", SqlDbType.VarChar, 50, "ProdStation")
            CLMasterSQLCommand.Parameters.Add("@LastChangedBy", SqlDbType.VarChar, 50, "LastChangedBy")
            CLMasterSQLCommand.Parameters.Add("@SupplyType", SqlDbType.VarChar, 20, "SupplyType")
            CLMasterSQLCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")

            CLMasterSQLCommand1 = New SqlClient.SqlCommand("INSERT INTO T_PDTOTemplateItem (TOTemplate,Item,Material,Manufacturer,MPN,Qty,MaterialRevision) values (@TOTemplate,@Item,@Material,@Manufacturer,@MPN,@Qty,@MaterialRevision)", myConn)

            CLMasterSQLCommand1.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand1.Parameters.Add("@Item", SqlDbType.Int, 4, "Item")
            CLMasterSQLCommand1.Parameters.Add("@Material", SqlDbType.VarChar, 50, "Material")
            CLMasterSQLCommand1.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 100, "Manufacturer")
            CLMasterSQLCommand1.Parameters.Add("@MPN", SqlDbType.VarChar, 100, "MPN")
            CLMasterSQLCommand1.Parameters.Add("@Qty", SqlDbType.Decimal, 13, "Qty")
            CLMasterSQLCommand1.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 50, "MaterialRevision")

            CLMasterSQLCommand.Parameters("@TOTemplate").Value = TemplateStr.Template
            CLMasterSQLCommand.Parameters("@Description").Value = TemplateStr.Desc
            CLMasterSQLCommand.Parameters("@ProdFloor").Value = Header.ProdFloor
            'CLMasterSQLCommand.Parameters("@ProdStation").Value = Header.ProdStation
            CLMasterSQLCommand.Parameters("@LastChangedBy").Value = ERPLoginData.User.ToUpper
            CLMasterSQLCommand.Parameters("@SupplyType").Value = Header.SupplyType
            CLMasterSQLCommand.Parameters("@ReasonCode").Value = Header.ReasonCode

            CLMasterSQLCommand.Parameters("@Product").Value = TOMydataSet.Tables("POItems").Rows(0)("Product")
            CLMasterSQLCommand.CommandType = CommandType.Text
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            ra = CLMasterSQLCommand.ExecuteNonQuery()

            For j = 0 To TOMydataSet.Tables("ItemDetails").Rows.Count - 1
                CLMasterSQLCommand1.Parameters("@TOTemplate").Value = TemplateStr.Template
                CLMasterSQLCommand1.Parameters("@Item").Value = j + 1
                CLMasterSQLCommand1.Parameters("@Material").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Material")
                CLMasterSQLCommand1.Parameters("@MPN").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("ManufacturerPN")
                CLMasterSQLCommand1.Parameters("@Manufacturer").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Manufacturer")
                CLMasterSQLCommand1.Parameters("@Qty").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("PerQty")
                CLMasterSQLCommand1.Parameters("@MaterialRevision").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Revision")
                CLMasterSQLCommand1.CommandType = CommandType.Text
                CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand1.ExecuteNonQuery()
            Next
            myConn.Close()

            Save_Template = "okay"
        Catch ex As Exception
            ErrorLogging("MMC-Save_Template", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Save_Template = "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetTemplateItem(ByVal Template As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                sqlstr = String.Format("select Material as 'Material',MaterialRevision as 'Revision',MPN as 'ManufacturerPN',Manufacturer as 'Manufacturer', Qty as 'PerQty' from T_PDTOTemplateItem with (nolock) where TOTemplate = '{0}'", Template)
                Return da.ExecuteDataSet(sqlstr, "TemplateItems")
            Catch ex As Exception
                Throw ex
                ErrorLogging("MMC-GetTemplateItem", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetTOItem(ByVal TOID As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim TemplateItem As DataTable
        Dim POItem, POHeader As DataTable
        Dim myDataColumn As DataColumn
        Dim Status As String

        GetTOItem = New DataSet
        TemplateItem = New Data.DataTable("TO")

        myDataColumn = New Data.DataColumn("Material", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.Decimal"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.Decimal"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("PickedQty", System.Type.GetType("System.Decimal"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("OpenQty", System.Type.GetType("System.Decimal"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Forming", System.Type.GetType("System.String"))
        TemplateItem.Columns.Add(myDataColumn)

        GetTOItem.Tables.Add(TemplateItem)

        POItem = New Data.DataTable("POItem")

        myDataColumn = New Data.DataColumn("PONo", System.Type.GetType("System.String"))
        POItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Style", System.Type.GetType("System.String"))
        POItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("POQty", System.Type.GetType("System.String"))
        POItem.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("IssueQty", System.Type.GetType("System.String"))
        POItem.Columns.Add(myDataColumn)
        GetTOItem.Tables.Add(POItem)

        POHeader = New Data.DataTable("POHeader")

        myDataColumn = New Data.DataColumn("ProdFloor", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("ProdLine", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("SupplyType", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("ETA", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("ReasonCode", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
        POHeader.Columns.Add(myDataColumn)

        GetTOItem.Tables.Add(POHeader)
        Dim Temp As Double
        Dim LabelIDsRow As Data.DataRow
        Dim LabelIDsRow1, LabelIDsRow2 As Data.DataRow
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1, CLMasterSQLCommand2 As SqlClient.SqlCommand

        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim objReader1, objReader2 As SqlClient.SqlDataReader
        CLMasterSQLCommand = New SqlClient.SqlCommand("select Material,MaterialRevision, MPN,Manufacturer,ReqQty,PickedQty, Description, Forming from T_PDTOItem with (nolock) INNER JOIN  T_SysLOV with (nolock) ON T_PDTOItem.Status = T_SysLOV.ProcessID where PDTO = @PDTO", myConn)
        CLMasterSQLCommand.Parameters.Add("@PDTO", SqlDbType.NVarChar, 50, "PDTO").Value = TOID
        'CLMasterSQLCommand.Parameters("@PDTO").Value = TOID

        CLMasterSQLCommand1 = New SqlClient.SqlCommand("select PO, Product,BuildQty from T_PDTO_PO with (nolock) where PDTO = @PDTO", myConn)
        CLMasterSQLCommand1.Parameters.Add("@PDTO", SqlDbType.NVarChar, 50, "PDTO")
        CLMasterSQLCommand1.Parameters("@PDTO").Value = TOID

        CLMasterSQLCommand2 = New SqlClient.SqlCommand("select ProdFloor,ProdLine,SupplyType,ETA,ReasonCode,Status from T_PDTOHeader with (nolock) where PDTO = @PDTO", myConn)
        CLMasterSQLCommand2.Parameters.Add("@PDTO", SqlDbType.NVarChar, 50, "PDTO")
        CLMasterSQLCommand2.Parameters("@PDTO").Value = TOID

        Dim i As Integer
        Dim sum As Double
        Dim comp1 As String

        Try
            myConn.Open()
            i = 0
            CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
            objReader1 = CLMasterSQLCommand1.ExecuteReader()
            While objReader1.Read()

                LabelIDsRow1 = GetTOItem.Tables("POItem").NewRow()
                If Not objReader1.GetValue(0) Is DBNull.Value Then LabelIDsRow1("PONo") = objReader1.GetValue(0)
                If Not objReader1.GetValue(1) Is DBNull.Value Then LabelIDsRow1("Style") = objReader1.GetValue(1)
                If Not objReader1.GetValue(2) Is DBNull.Value Then
                    If i > 0 Then
                        If comp1 = objReader1.GetValue(1) Then
                            sum = Temp + objReader1.GetValue(2)
                        Else
                            sum = objReader1.GetValue(2)
                        End If
                        LabelIDsRow1("IssueQty") = objReader1.GetValue(2)
                    Else
                        comp1 = objReader1.GetValue(1)
                        LabelIDsRow1("IssueQty") = objReader1.GetValue(2)
                        Temp = objReader1.GetValue(2)
                        sum = objReader1.GetValue(2)
                    End If
                End If
                If Me.GetDJInfoFromERP(ERPLoginData, LabelIDsRow1("PONo")).Tables(0).Rows.Count > 0 Then
                    LabelIDsRow1("POQty") = Me.GetDJInfoFromERP(ERPLoginData, LabelIDsRow1("PONo")).Tables(0).Rows(0).Item(7)
                Else
                    LabelIDsRow1("POQty") = 0
                End If

                GetTOItem.Tables("POItem").Rows.Add(LabelIDsRow1)
                i = i + 1

            End While
            objReader1.Close()
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            objReader2 = CLMasterSQLCommand2.ExecuteReader()
            While objReader2.Read()
                LabelIDsRow2 = GetTOItem.Tables("POHeader").NewRow()
                If Not objReader2.GetValue(0) Is DBNull.Value Then LabelIDsRow2("ProdFloor") = objReader2.GetValue(0)
                If Not objReader2.GetValue(1) Is DBNull.Value Then LabelIDsRow2("ProdLine") = objReader2.GetValue(1)
                If Not objReader2.GetValue(2) Is DBNull.Value Then LabelIDsRow2("SupplyType") = objReader2.GetValue(2)
                If Not objReader2.GetValue(3) Is DBNull.Value Then LabelIDsRow2("ETA") = objReader2.GetValue(3)
                If Not objReader2.GetValue(4) Is DBNull.Value Then LabelIDsRow2("ReasonCode") = objReader2.GetValue(4)
                If Not objReader2.GetValue(5) Is DBNull.Value Then LabelIDsRow2("Status") = objReader2.GetValue(5)
                GetTOItem.Tables("POHeader").Rows.Add(LabelIDsRow2)
            End While
            objReader2.Close()

            Status = GetTOItem.Tables("POHeader").Rows(0)("Status").ToString

            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                LabelIDsRow = GetTOItem.Tables("TO").NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then LabelIDsRow("Material") = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then LabelIDsRow("Revision") = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then LabelIDsRow("ManufacturerPN") = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then LabelIDsRow("Manufacturer") = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then LabelIDsRow("ReqQty") = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then LabelIDsRow("PickedQty") = objReader.GetValue(5)
                If Status = 0 Then
                    LabelIDsRow("Status") = "Closed"
                Else
                    If Not objReader.GetValue(6) Is DBNull.Value Then LabelIDsRow("Status") = objReader.GetValue(6)
                End If

                If LabelIDsRow("ReqQty") Is DBNull.Value Then
                    LabelIDsRow("ReqQty") = 1
                End If
                LabelIDsRow("PerQty") = LabelIDsRow("ReqQty") / sum
                If Not LabelIDsRow("PickedQty") Is DBNull.Value Then
                    LabelIDsRow("OpenQty") = CDec(LabelIDsRow("ReqQty")) - CDec(LabelIDsRow("PickedQty"))
                Else
                    LabelIDsRow("OpenQty") = LabelIDsRow("ReqQty")
                End If
                If Not objReader.GetValue(7) Is DBNull.Value Then LabelIDsRow("Forming") = objReader.GetValue(7)

                GetTOItem.Tables("TO").Rows.Add(LabelIDsRow)
            End While
            objReader.Close()

            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-GetTOItem", ERPLoginData.User.ToUpper, "TOID: " & TOID & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Search_Template(ByVal Template As String) As String
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        CLMasterSQLCommand = New SqlClient.SqlCommand("select TOTemplate, Description, LastChangedBy, LastChangedOn from T_PDTOTemplateHeader with (nolock) where TOTemplate = @TOTemplate", myConn)
        CLMasterSQLCommand.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
        CLMasterSQLCommand.Parameters("@TOTemplate").Value = Template
        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
        Try
            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then
                    Search_Template = "Okay" 'alreay existent
                Else
                    Search_Template = "Error"   'don't existent
                End If
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-Search_Template", "", "Template: " & Template & ", " & ex.Message & ex.Source, "E")
            Search_Template = "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Check_Template(ByVal Template As String, ByVal ERPLoginData As ERPLogin) As String
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        CLMasterSQLCommand = New SqlClient.SqlCommand("select TOTemplate, Description, LastChangedBy, LastChangedOn from T_PDTOTemplateHeader with (nolock) where TOTemplate = @TOTemplate", myConn)
        CLMasterSQLCommand.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
        CLMasterSQLCommand.Parameters("@TOTemplate").Value = Template
        CLMasterSQLCommand.CommandTimeout = TimeOut_M5
        Try
            myConn.Open()
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then
                    Check_Template = "Exist" 'alreay existent
                Else
                    Check_Template = "Blank"   'don't existent
                End If
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-Search_Template", "", "Template: " & Template & ", " & ex.Message & ex.Source, "E")
            Check_Template = "Blank"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Search_LockTO(ByVal TOVALUE As String) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SearchCommand As SqlClient.SqlCommand
        Dim objReader As SqlClient.SqlDataReader
        SearchCommand = New SqlClient.SqlCommand("select PDTO from T_PDTOHeader with (nolock) where PDTO = @PDTO and Status = 2", myConn)
        SearchCommand.Parameters.Add("@PDTO", SqlDbType.NVarChar, 50, "PDTO")
        SearchCommand.Parameters("@PDTO").Value = TOVALUE
        SearchCommand.CommandTimeout = TimeOut_M5
        Try
            myConn.Open()
            objReader = SearchCommand.ExecuteReader()
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then
                    Search_LockTO = True 'alreay existent
                Else
                    Search_LockTO = False   'don't existent
                End If
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-Search_LockTO", "", "TOVALUE: " & TOVALUE & ", " & ex.Message & ex.Source, "E")
            Search_LockTO = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function Search_OpenTO(ByVal TOVALUE As String) As Boolean
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim SearchCommand As SqlClient.SqlCommand
        Dim objReader As SqlClient.SqlDataReader
        SearchCommand = New SqlClient.SqlCommand("select PDTO from T_PDTOHeader with (nolock) where PDTO = @PDTO and Status in ('1','2')", myConn)
        SearchCommand.Parameters.Add("@PDTO", SqlDbType.NVarChar, 50, "PDTO")
        SearchCommand.Parameters("@PDTO").Value = TOVALUE

        Try
            myConn.Open()
            objReader = SearchCommand.ExecuteReader()
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then
                    Search_OpenTO = True 'alreay existent
                Else
                    Search_OpenTO = False   'don't existent
                End If
            End While
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-Search_OpenTO", "", "TOVALUE: " & TOVALUE & ", " & ex.Message & ex.Source, "E")
            Search_OpenTO = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function Unlock_PickOrder(ByVal PickOrder As String, ByVal erplogindata As ERPLogin) As Boolean
        Dim UpdateCommand1 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Try
            myConn.Open()
            UpdateCommand1 = New SqlClient.SqlCommand("update T_PDTOHeader set Status=1,ChangedOn=getdate(),ChangedBy=@Changedby where PDTO=@PDTO", myConn)
            UpdateCommand1.Parameters.Add("@ChangedBy", SqlDbType.VarChar, 50, "ChangedBy")
            UpdateCommand1.Parameters.Add("@PDTO", SqlDbType.VarChar, 50, "PDTO")
            UpdateCommand1.Parameters("@Changedby").Value = erplogindata.User
            UpdateCommand1.Parameters("@PDTO").Value = PickOrder
            UpdateCommand1.CommandType = CommandType.Text
            UpdateCommand1.CommandTimeout = TimeOut_M5
            ra = UpdateCommand1.ExecuteNonQuery()
            myConn.Close()

            Unlock_PickOrder = True
        Catch ex As Exception
            ErrorLogging("MMC-Unlock_PickOrder", erplogindata.User, "PickOrder: " & PickOrder & ", " & ex.Message & ex.Source, "E")
            Unlock_PickOrder = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function ClosePickOrder(ByVal mydataset As DataSet, ByVal PickOrder As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Dim UpdateCommand As SqlClient.SqlCommand
        Dim UpdateCommand1 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim i As Integer
        Dim Remark As String = DateTime.Now.ToString & " by User: " & ERPLoginData.User.ToUpper & " with LastTransaction: ClosePickOrder"

        Try
            myConn.Open()
            'For i = 0 To mydataset.Tables(1).Rows.Count - 1

            '    UpdateCommand = New SqlClient.SqlCommand("update T_PDTOItem set Status=0,Remark=@Remark where Material=@Material", myConn)
            '    UpdateCommand.Parameters.Add("@Remark", SqlDbType.VarChar, 100, "Remark")
            '    UpdateCommand.Parameters.Add("@Material", SqlDbType.VarChar, 50, "Material")
            '    UpdateCommand.Parameters("@Remark").Value = Remark
            '    UpdateCommand.Parameters("@Material").Value = mydataset.Tables(1).Rows(i)(0)
            '    UpdateCommand.CommandType = CommandType.Text
            '    ra = UpdateCommand.ExecuteNonQuery()
            'Next
            UpdateCommand1 = New SqlClient.SqlCommand("update T_PDTOHeader set Status=0,ChangedOn=getdate(),ChangedBy=@Changedby where PDTO=@PDTO", myConn)
            UpdateCommand1.Parameters.Add("@Changedby", SqlDbType.VarChar, 50, "Changedby")
            UpdateCommand1.Parameters.Add("@PDTO", SqlDbType.VarChar, 50, "PDTO")
            UpdateCommand1.Parameters("@Changedby").Value = ERPLoginData.User
            UpdateCommand1.Parameters("@PDTO").Value = PickOrder
            UpdateCommand1.CommandTimeout = TimeOut_M5
            UpdateCommand1.CommandType = CommandType.Text
            ra = UpdateCommand1.ExecuteNonQuery()
            myConn.Close()

            ClosePickOrder = True
            'ClosePickOrder = False
        Catch ex As Exception
            ErrorLogging("MMC-ClosePickOrder", ERPLoginData.User, "PickOrder: " & PickOrder & ", " & ex.Message & ex.Source, "E")
            ClosePickOrder = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function Get_PickOrder(ByVal header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim TOIDs As DataTable
        Dim myDataColumn As DataColumn

        TOIDs = New Data.DataTable("TO")
        myDataColumn = New Data.DataColumn("TOInfo", System.Type.GetType("System.String"))
        TOIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LastChangedBy", System.Type.GetType("System.String"))
        TOIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("LastChangedOn", System.Type.GetType("System.String"))
        TOIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("DiscreteJob", System.Type.GetType("System.String"))
        TOIDs.Columns.Add(myDataColumn)
        myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
        TOIDs.Columns.Add(myDataColumn)
        ReturnTemplate.Tables.Add(TOIDs)

        Dim LabelIDsRow As Data.DataRow
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim objReader As SqlClient.SqlDataReader
        Dim temp As String

        If header.ProdLine = "" Then
            If header.PO = "" Then
                temp = "SELECT T_PDTOHeader.PDTO, CreatedBy, CreatedOn, ChangedBy, ChangedOn, PO, Description FROM T_PDTO_PO with (nolock) INNER JOIN T_PDTOHeader with (nolock) ON T_PDTO_PO.PDTO = T_PDTOHeader.PDTO INNER JOIN  T_SysLOV ON T_PDTOHeader.Status = T_SysLOV.ProcessID WHERE ProdFloor = @ProdFloor and T_PDTOHeader.OrgCode = @OrgCode AND Name = 'Pick Order Status'"
            Else
                temp = "SELECT T_PDTOHeader.PDTO, CreatedBy, CreatedOn, ChangedBy, ChangedOn, PO, Description FROM T_PDTO_PO with (nolock) INNER JOIN T_PDTOHeader with (nolock) ON T_PDTO_PO.PDTO = T_PDTOHeader.PDTO INNER JOIN  T_SysLOV ON T_PDTOHeader.Status = T_SysLOV.ProcessID WHERE PO = @PO and ProdFloor = @ProdFloor and T_PDTOHeader.OrgCode = @OrgCode AND Name = 'Pick Order Status'"
            End If
        Else
            If header.PO = "" Then
                temp = "SELECT T_PDTOHeader.PDTO, CreatedBy, CreatedOn, ChangedBy, ChangedOn, PO, Description FROM T_PDTO_PO with (nolock) INNER JOIN T_PDTOHeader with (nolock) ON T_PDTO_PO.PDTO = T_PDTOHeader.PDTO INNER JOIN  T_SysLOV ON T_PDTOHeader.Status = T_SysLOV.ProcessID  WHERE ProdFloor = @ProdFloor and ProdLine = @ProdLine and T_PDTOHeader.OrgCode = @OrgCode AND Name = 'Pick Order Status'"
            Else
                temp = "SELECT T_PDTOHeader.PDTO, CreatedBy, CreatedOn, ChangedBy, ChangedOn, PO, Description FROM T_PDTO_PO with (nolock) INNER JOIN T_PDTOHeader with (nolock) ON T_PDTO_PO.PDTO = T_PDTOHeader.PDTO INNER JOIN  T_SysLOV ON T_PDTOHeader.Status = T_SysLOV.ProcessID  WHERE PO = @PO and ProdFloor = @ProdFloor and ProdLine = @ProdLine and T_PDTOHeader.OrgCode = @OrgCode AND Name = 'Pick Order Status'"
            End If
        End If
        'If header.ReasonCode <> "" Then
        '    temp = temp & " And ReasonCode = @ReasonCode"
        'End If
        CLMasterSQLCommand = New SqlClient.SqlCommand(temp, myConn)
        CLMasterSQLCommand.Parameters.Add("@ProdFloor", SqlDbType.VarChar, 50, "ProdFloor")
        CLMasterSQLCommand.Parameters("@ProdFloor").Value = header.ProdFloor
        If header.ProdLine <> "" Then
            CLMasterSQLCommand.Parameters.Add("@ProdLine", SqlDbType.VarChar, 50, "ProdLine")
            CLMasterSQLCommand.Parameters("@ProdLine").Value = header.ProdLine
        End If
        If header.PO <> "" Then
            CLMasterSQLCommand.Parameters.Add("@PO", SqlDbType.VarChar, 50, "PO")
            CLMasterSQLCommand.Parameters("@PO").Value = header.PO
        End If
        CLMasterSQLCommand.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")
        CLMasterSQLCommand.Parameters("@OrgCode").Value = ERPLoginData.OrgCode
        'CLMasterSQLCommand.Parameters.Add("@SupplyType", SqlDbType.VarChar, 20, "SupplyType")
        'CLMasterSQLCommand.Parameters("@SupplyType").Value = header.SupplyType
        'CLMasterSQLCommand.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
        'CLMasterSQLCommand.Parameters("@ReasonCode").Value = header.ReasonCode

        Try
            myConn.Open()
            CLMasterSQLCommand.CommandTimeout = TimeOut_M5
            objReader = CLMasterSQLCommand.ExecuteReader()
            While objReader.Read()
                'Fill LabelIDs table
                LabelIDsRow = ReturnTemplate.Tables("TO").NewRow()
                If Not objReader.GetValue(0) Is DBNull.Value Then LabelIDsRow("TOInfo") = objReader.GetValue(0)
                If Not objReader.GetValue(3) Is DBNull.Value Then
                    LabelIDsRow("LastChangedBy") = objReader.GetValue(3)
                ElseIf Not objReader.GetValue(1) Is DBNull.Value Then
                    LabelIDsRow("LastChangedBy") = objReader.GetValue(1)
                End If
                If Not objReader.GetValue(4) Is DBNull.Value Then
                    LabelIDsRow("LastChangedOn") = objReader.GetValue(4)
                ElseIf Not objReader.GetValue(2) Is DBNull.Value Then
                    LabelIDsRow("LastChangedOn") = objReader.GetValue(2)
                End If
                If Not objReader.GetValue(5) Is DBNull.Value Then LabelIDsRow("DiscreteJob") = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then LabelIDsRow("Status") = objReader.GetValue(6)

                ReturnTemplate.Tables("TO").Rows.Add(LabelIDsRow)
            End While
            myConn.Close()

            If header.Status = "All" Then

            ElseIf header.Status = "Open" Then
                Dim DR_Blank() As DataRow = Nothing
                DR_Blank = ReturnTemplate.Tables("TO").Select("status = 'Closed'")
                If DR_Blank.Length > 0 Then
                    For Each drblank As DataRow In DR_Blank
                        ReturnTemplate.Tables("TO").Rows.Remove(drblank)
                    Next
                End If
            ElseIf header.Status = "Closed" Then
                Dim DR_Blank() As DataRow = Nothing
                DR_Blank = ReturnTemplate.Tables("TO").Select("status = 'Open'")
                If DR_Blank.Length > 0 Then
                    For Each drblank As DataRow In DR_Blank
                        ReturnTemplate.Tables("TO").Rows.Remove(drblank)
                    Next
                End If
            End If

        Catch ex As Exception
            ErrorLogging("MMC-GetPickOrder", "", ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return ReturnTemplate
    End Function

    Public Function GetTemplate(ByVal Header As ProdPickingStructure) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                sqlstr = String.Format("select TOTemplate as 'Template', Description as 'Desc', LastChangedBy, LastChangedOn from T_PDTOTemplateHeader with (nolock) where ProdFloor = '{0}'", Header.ProdFloor)
                Return da.ExecuteDataSet(sqlstr, "Templates")
            Catch ex As Exception
                Throw ex
                ErrorLogging("MMC-GetTemplate", "", "ProdFloor: " & Header.ProdFloor & ", " & ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetDJInfo(ByVal ID As String, ByVal ERPLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                If Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    'Dim i As Integer
                    'i = da.ExecuteScalar(String.Format("select count(distinct PurOrdNo) from T_CLMaster where BoxID = '{0}'", ID))
                    Dim ds As New DataSet
                    ds = da.ExecuteDataSet(String.Format("select count(distinct PurOrdNo) as c1,count(distinct MaterialNo) as c2,count(distinct MaterialRevision) as c3,count(distinct OrgCode) as c4,count(distinct SLOC) as c5,count(distinct StorageBin) as c6,count(distinct LotNo) as c7,count(distinct StatusCode) as c8 from T_CLMaster with (nolock) where BoxID = '{0}'", ID), "IDInfo")
                    If ds.Tables(0).Rows(0)(0) = 1 AndAlso ds.Tables(0).Rows(0)(1) = 1 AndAlso ds.Tables(0).Rows(0)(2) = 1 AndAlso ds.Tables(0).Rows(0)(3) = 1 AndAlso ds.Tables(0).Rows(0)(4) = 1 AndAlso ds.Tables(0).Rows(0)(5) = 1 AndAlso ds.Tables(0).Rows(0)(6) = 1 AndAlso ds.Tables(0).Rows(0)(7) = 1 Then
                        Dim status As String
                        status = da.ExecuteScalar(String.Format("select StatusCode from T_CLMaster where BoxID = '{0}'", ID))
                        If status = "1" Then
                            Dim sqlstr As String
                            sqlstr = String.Format("select PurOrdNo as 'DJ',LotNo as 'LotNo',MaterialNo as 'Mat',MaterialRevision as 'Rev',OrgCode as 'OrgCode',SUM(QtyBaseUOM) as 'Qty', BaseUOM as 'UOM', SLOC as 'SubInventory',StorageBin as 'Locator' from T_CLMaster with (nolock) where BoxID = '{0}' and StatusCode = 1 and PurOrdNo <>'' group by PurOrdNo,LotNo,MaterialNo,MaterialRevision,OrgCode,BaseUOM,SLOC,StorageBin", ID)
                            Return da.ExecuteDataSet(sqlstr, "DJInfo")
                        End If
                    End If
                Else
                    Dim sqlstr As String
                    sqlstr = String.Format("select PurOrdNo as 'DJ',LotNo as 'LotNo',MaterialNo as 'Mat',MaterialRevision as 'Rev',OrgCode as 'OrgCode',QtyBaseUOM as 'Qty', BaseUOM as 'UOM', SLOC as 'SubInventory',StorageBin as 'Locator' from T_CLMaster with (nolock) where CLID = '{0}' and StatusCode = 1 and PurOrdNo <>''", ID)
                    Return da.ExecuteDataSet(sqlstr, "DJInfo")
                End If
            Catch ex As Exception
                ErrorLogging("MMC-GetDJInfo", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Throw ex
            End Try
        End Using
    End Function

    Public Function Update_CLID(ByVal ID As String, ByVal TransType As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                If Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    sqlstr = String.Format("Update T_CLMaster set ChangedOn=getdate(), ChangedBy='{0}', StatusCode=0,LastTransaction='{1}' where BoxID='{2}' ", ERPLoginData.User, TransType, ID)
                Else
                    sqlstr = String.Format("Update T_CLMaster set ChangedOn=getdate(), ChangedBy='{0}', StatusCode=0,LastTransaction='{1}' where CLID='{2}' ", ERPLoginData.User, TransType, ID)
                End If
                da.ExecuteNonQuery(sqlstr)
                Return True
            Catch ex As Exception
                ErrorLogging(TransType & "-UpdateCLID", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function UpdateCLID(ByVal ID As String, ByVal ERPLoginData As ERPLogin) As Boolean
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                sqlstr = String.Format("Update T_CLMaster set ChangedOn = getdate(), ChangedBy = '{0}', StatusCode=0,LastTransaction='UpdateCLID' where CLID='{1}' ", ERPLoginData.User, ID)
                da.ExecuteNonQuery(sqlstr)
                Return True
            Catch ex As Exception
                ErrorLogging("MMC-UpdateCLID", "", ex.Message & ex.Source, "E")
                Return False
            End Try
        End Using
    End Function

    Public Function Post_DJ_Reversal(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal orgcode As String, ByVal return_qty As Integer, ByVal uom As String, ByVal SubInventory As String, ByVal Locator As String, ByVal CLID As String) As dj_response
        Using da As DataAccess = GetDataAccess()
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator)
            Dim aa As OracleString
            Dim bb As Int32
            'Dim rev As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()
            Try
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_return"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'orgcode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_return_qty", OracleType.Number, 50).Value = return_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                'Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = rev
                Oda.Parameters.Add("o_subinventory", OracleType.VarChar, 150).Value = SubInventory.Trim
                Oda.Parameters.Add("o_locator", OracleType.VarChar, 100).Value = Locator.Trim
                Oda.Parameters("o_subinventory").Direction = ParameterDirection.InputOutput
                Oda.Parameters("o_locator").Direction = ParameterDirection.InputOutput
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output
                bb = Oda.ExecuteOracleNonQuery(aa)

                Post_DJ_Reversal.flag = Oda.Parameters("o_success_flag").Value.ToString()
                Post_DJ_Reversal.errormsg = Oda.Parameters("o_error_msgg").Value.ToString()
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()

                    Dim UpdateCLID As Boolean
                    UpdateCLID = Update_CLID(CLID, "DJ_Reversal", ERPLoginData)
                    If UpdateCLID = False Then
                        Post_DJ_Reversal.flag = "N"
                        Post_DJ_Reversal.errormsg = "eTrace update error"
                    End If
                Else
                    Oda.Transaction.Rollback()
                End If

                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return Post_DJ_Reversal
            Catch ex As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("MaterialReplenishment-DJ_Reversal", "", "DJ: " & DJ & ", " & ex.Message & ex.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    Oda.Connection.Dispose()
                End If
                Post_DJ_Reversal.flag = "N"
                Post_DJ_Reversal.errormsg = ex.Message.ToString
                Return Post_DJ_Reversal
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try

        End Using
    End Function

    Public Function DJ_Reversal(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal orgcode As String, ByVal return_qty As Integer, ByVal uom As String, ByVal SubInventory As String, ByVal Locator As String) As dj_response
        Using da As DataAccess = GetDataAccess()
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator)
            Dim aa As OracleString
            Dim bb As Int32
            'Dim rev As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()
            Try

                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_return"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'orgcode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_return_qty", OracleType.Number, 50).Value = return_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                'Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = rev
                Oda.Parameters.Add("o_subinventory", OracleType.VarChar, 150).Value = SubInventory.Trim
                Oda.Parameters.Add("o_locator", OracleType.VarChar, 100).Value = Locator.Trim
                Oda.Parameters("o_subinventory").Direction = ParameterDirection.InputOutput
                Oda.Parameters("o_locator").Direction = ParameterDirection.InputOutput
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output
                bb = Oda.ExecuteOracleNonQuery(aa)
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()
                Else
                    Oda.Transaction.Rollback()
                End If
                DJ_Reversal.flag = Oda.Parameters("o_success_flag").Value.ToString()
                DJ_Reversal.errormsg = Oda.Parameters("o_error_msgg").Value.ToString()
                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                'Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return DJ_Reversal
            Catch oe As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("MaterialReplenishment-DJ_Reversal", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    ' Oda.Connection.Dispose()
                End If
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Check_PickList(ByVal mydataset As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As String
        Try
            Dim i As Integer
            Dim Validate_Rev As ItemRevList

            For i = 0 To mydataset.Tables(1).Rows.Count - 1
                If Not mydataset.Tables(1).Rows(i)("Material") Is DBNull.Value Then
                    If FixNull(mydataset.Tables(1).Rows(i)("Revision")) <> "" Then
                        mydataset.Tables(1).Rows(i)("Revision") = UCase(FixNull(mydataset.Tables(1).Rows(i)("Revision")))
                    End If
                    Validate_Rev = ValidateItemRevision(ERPLoginData, mydataset.Tables(1).Rows(i)("Material"), FixNull(mydataset.Tables(1).Rows(i)("Revision")), "Pick Order Creation")
                    If Validate_Rev.Flag = "Y" AndAlso FixNull(mydataset.Tables(1).Rows(i)("Revision")) = "" OrElse Validate_Rev.Flag = "N" AndAlso FixNull(mydataset.Tables(1).Rows(i)("Revision")) <> "" OrElse Validate_Rev.Flag <> "Y" AndAlso Validate_Rev.Flag <> "N" Then
                        If InStr(Check_PickList, "Revision") < 1 Then
                            If FixNull(Check_PickList) = "" Then
                                Check_PickList = "Revision"
                            Else
                                Check_PickList = Check_PickList & " & " & "Revision"
                            End If
                        End If
                        '    Check_PickList = "Revision"
                        '    'Me.ShowMessage("^MMC-72@", eTraceUI.eTraceMessageBar.MessageTypes.Abort)
                        '    Exit Function
                    End If
                End If

                If Not mydataset.Tables(1).Rows(i)("Material") Is DBNull.Value AndAlso Not mydataset.Tables(1).Rows(i)("PerQty") Is DBNull.Value Then
                    If Not mydataset.Tables(1).Rows(i)("Shortage") Is DBNull.Value AndAlso mydataset.Tables(1).Rows(i)("Shortage") > 0 Then
                        If InStr(Check_PickList, "Shortage") < 1 Then
                            If FixNull(Check_PickList) = "" Then
                                Check_PickList = "Shortage"
                            Else
                                Check_PickList = Check_PickList & " & " & "Shortage"
                            End If
                        End If
                        'Check_PickList = "Shortage"
                        'Exit Function
                    End If
                End If

                If FixNull(mydataset.Tables(1).Rows(i)("Material")) = "" OrElse FixNull(mydataset.Tables(1).Rows(i)("PerQty")) = "" OrElse (UCase(Header.SupplyType) = "MSB" AndAlso (FixNull(mydataset.Tables(1).Rows(i)("ManufacturerPN")) = "" OrElse FixNull(mydataset.Tables(1).Rows(i)("Manufacturer")) = "")) Then
                    If InStr(Check_PickList, "Record") < 1 Then
                        If FixNull(Check_PickList) = "" Then
                            Check_PickList = "Record"
                        Else
                            Check_PickList = Check_PickList & " & " & "Record"
                        End If
                    End If
                    'Check_PickList = "Record"
                    ''Me.ShowMessage("^MMC-25@", eTraceUI.eTraceMessageBar.MessageTypes.Abort)
                    'Exit Function
                End If
            Next
        Catch ex As Exception
            ErrorLogging("PickOrderCreation Check_PickList", "", ex.Message & ex.Source, "E")
        End Try
    End Function

    Public Function PostProdPicking(ByVal TOMydataSet As DataSet, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As ProdPickingResponse

        Dim i As Integer
        Dim j, k As Integer

        Dim myCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1 As SqlClient.SqlCommand
        Dim CLMasterSQLCommand2 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim NextID As String
        Dim DestInv As DataSet
        Dim HeaderRs As ProdPickingStructure

        myCommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.Parameters.AddWithValue("@NextNo", "")
        myCommand.Parameters(0).SqlDbType = SqlDbType.VarChar
        myCommand.Parameters(0).Size = 20
        myCommand.Parameters(0).Direction = ParameterDirection.Output
        myCommand.Parameters.AddWithValue("@TypeID", "PTOID")
        myCommand.CommandTimeout = TimeOut_M5
        Try
            myConn.Open()

            'Try up to 5 times when failed getting next id
            NextID = ""
            k = 0
            While (k < 5 And NextID = "")
                Try
                    ra = myCommand.ExecuteNonQuery()
                    NextID = myCommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("MMC-PostProdPicking", "Deadlocked? " & Str(k), "Failed getting next ID; PO: " & Header.PO & ", " & ex.Message & ex.Source, "E")
                End Try
            End While

            If NextID = "" Then
                PostProdPicking.ErrMessage = "Error"
                Exit Function
            End If

            CLMasterSQLCommand = New SqlClient.SqlCommand("INSERT INTO T_PDTO_PO (PDTO,PO,Product,BuildQty) values (@PDTO,@PO,@Product,@BuildQty)", myConn)

            CLMasterSQLCommand.Parameters.Add("@PDTO", SqlDbType.VarChar, 20, "PDTO")
            CLMasterSQLCommand.Parameters.Add("@PO", SqlDbType.VarChar, 50, "PO")
            CLMasterSQLCommand.Parameters.Add("@Product", SqlDbType.VarChar, 50, "Product")
            CLMasterSQLCommand.Parameters.Add("@BuildQty", SqlDbType.Int, 4, "BuildQty")
            If UCase(Header.SupplyType) = "MSB" OrElse UCase(Header.SupplyType) = "PULL" OrElse UCase(Header.SupplyType) = "PRE-WORK" OrElse UCase(Header.SupplyType) = "FLOORSTOCK" OrElse UCase(Header.SupplyType) = "OSP" Then
                CLMasterSQLCommand1 = New SqlClient.SqlCommand("INSERT INTO T_PDTOHeader (PDTO,ProdFloor,ProdLine,ETA,MultiPOs,CreatedBy,CreatedOn,Status,SupplyType,DestSubInv,DestLocator,ReasonCode,OrgCode) values (@PDTO,@ProdFloor,@ProdLine,@ETA,@MultiPOs,@CreatedBy,getdate(),1,@SupplyType,@DestSubInv,@DestLocator,@ReasonCode,@OrgCode)", myConn)
                CLMasterSQLCommand1.Parameters.Add("@DestSubInv", SqlDbType.VarChar, 50, "DestSubInv")
                CLMasterSQLCommand1.Parameters.Add("@DestLocator", SqlDbType.VarChar, 50, "DestLocator")
                'CLMasterSQLCommand1.Parameters("@DestSubInv").Value = GetOrderInfoFromERP(ERPLoginData, TOMydataSet.Tables("POItems").Rows(0)("PONo")).Tables(0).Rows(0)(5)

                'If FixNull(Header.CheckSubInv) = "" orelse fixnull(Header.CheckLocator) = "" Then
                'HeaderRs = Get_DestSubLoc(TOMydataSet.Tables("POItems").Rows(0)("DJNo"), ERPLoginData, Header)
                'CLMasterSQLCommand1.Parameters("@DestSubInv").Value = HeaderRs.CheckSubInv
                'CLMasterSQLCommand1.Parameters("@DestLocator").Value = HeaderRs.CheckLocator
                'Else
                'CLMasterSQLCommand1.Parameters("@DestSubInv").Value = Header.CheckSubInv
                'CLMasterSQLCommand1.Parameters("@DestLocator").Value = Header.CheckLocator
                'End If

                If (UCase(Header.SupplyType) = "MSB" OrElse UCase(Header.SupplyType) = "PULL" OrElse UCase(Header.SupplyType) = "PRE-WORK") AndAlso FixNull(Header.CheckSubInv) = "" Then
                    DestInv = New DataSet
                    DestInv = Getrelease_lines(ERPLoginData, TOMydataSet.Tables("POItems").Rows(0)("DJNo"))
                    If Not DestInv.Tables("Pickorder2") Is Nothing AndAlso DestInv.Tables("Pickorder2").Rows.Count > 0 Then
                        CLMasterSQLCommand1.Parameters("@DestSubInv").Value = DestInv.Tables(0).Rows(0)(0)
                        CLMasterSQLCommand1.Parameters("@DestLocator").Value = DestInv.Tables(0).Rows(0)(1)
                    Else
                        ErrorLogging("MMC-PostProdPicking", ERPLoginData.User.ToUpper, "No Dest.Subinv available for Pick Order", "I")
                        PostProdPicking.ErrMessage = "Error:No Dest.Subinv available for Pick Order"
                        Exit Function
                    End If
                ElseIf (UCase(Header.SupplyType) = "MSB" OrElse UCase(Header.SupplyType) = "PULL" OrElse UCase(Header.SupplyType) = "PRE-WORK") AndAlso FixNull(Header.CheckSubInv) <> "" Then
                    CLMasterSQLCommand1.Parameters("@DestSubInv").Value = Header.CheckSubInv
                    CLMasterSQLCommand1.Parameters("@DestLocator").Value = Header.CheckLocator
                ElseIf UCase(Header.SupplyType) = "FLOORSTOCK" OrElse UCase(Header.SupplyType) = "OSP" Then
                    CLMasterSQLCommand1.Parameters("@DestSubInv").Value = Header.CheckSubInv
                    CLMasterSQLCommand1.Parameters("@DestLocator").Value = Header.CheckLocator
                End If
            Else
                CLMasterSQLCommand1 = New SqlClient.SqlCommand("INSERT INTO T_PDTOHeader (PDTO,ProdFloor,ProdLine,ETA,MultiPOs,CreatedBy,CreatedOn,Status,SupplyType,ReasonCode,OrgCode) values (@PDTO,@ProdFloor,@ProdLine,@ETA,@MultiPOs,@CreatedBy,getdate(),1,@SupplyType,@ReasonCode,@OrgCode)", myConn)
            End If
            CLMasterSQLCommand1.Parameters.Add("@PDTO", SqlDbType.VarChar, 20, "PDTO")
            CLMasterSQLCommand1.Parameters.Add("@ProdFloor", SqlDbType.VarChar, 50, "ProdFloor")
            CLMasterSQLCommand1.Parameters.Add("@ProdLine", SqlDbType.VarChar, 50, "ProdLine")
            CLMasterSQLCommand1.Parameters.Add("@ETA", SqlDbType.DateTime, 20, "ETA")
            CLMasterSQLCommand1.Parameters.Add("@MultiPOs", SqlDbType.Int, 4, "MultiPOs")
            CLMasterSQLCommand1.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50, "CreatedBy")
            CLMasterSQLCommand1.Parameters.Add("@SupplyType", SqlDbType.VarChar, 20, "SupplyType")
            CLMasterSQLCommand1.Parameters.Add("@ReasonCode", SqlDbType.VarChar, 50, "ReasonCode")
            CLMasterSQLCommand1.Parameters.Add("@OrgCode", SqlDbType.VarChar, 50, "OrgCode")

            CLMasterSQLCommand2 = New SqlClient.SqlCommand("INSERT INTO T_PDTOItem (PDTO,Item,Material,MaterialRevision,Manufacturer,MPN,ReqQty,PickedQty,Status,Forming) values (@PDTO,@Item,@Material,@MaterialRevision,@Manufacturer,@MPN,@ReqQty,@PickedQty,1,@Forming)", myConn)

            CLMasterSQLCommand2.Parameters.Add("@PDTO", SqlDbType.VarChar, 20, "PDTO")
            CLMasterSQLCommand2.Parameters.Add("@Item", SqlDbType.Int, 4, "Item")
            CLMasterSQLCommand2.Parameters.Add("@Material", SqlDbType.VarChar, 50, "Material")
            CLMasterSQLCommand2.Parameters.Add("@MaterialRevision", SqlDbType.VarChar, 50, "MaterialRevision")
            CLMasterSQLCommand2.Parameters.Add("@Manufacturer", SqlDbType.VarChar, 100, "Manufacturer")
            CLMasterSQLCommand2.Parameters.Add("@MPN", SqlDbType.VarChar, 100, "MPN")
            CLMasterSQLCommand2.Parameters.Add("@ReqQty", SqlDbType.Decimal, 13, "ReqQty")
            CLMasterSQLCommand2.Parameters.Add("@PickedQty", SqlDbType.Decimal, 13, "PickedQty")
            CLMasterSQLCommand2.Parameters.Add("@Forming", SqlDbType.VarChar, 10, "Forming")

            CLMasterSQLCommand.Parameters("@PDTO").Value = NextID
            CLMasterSQLCommand1.Parameters("@PDTO").Value = NextID
            CLMasterSQLCommand2.Parameters("@PDTO").Value = NextID

            CLMasterSQLCommand1.Parameters("@ProdFloor").Value = Header.ProdFloor
            CLMasterSQLCommand1.Parameters("@ProdLine").Value = Header.ProdLine
            CLMasterSQLCommand1.Parameters("@ETA").Value = Header.ETA
            CLMasterSQLCommand1.Parameters("@MultiPOs").Value = TOMydataSet.Tables("POItems").Rows.Count
            CLMasterSQLCommand1.Parameters("@CreatedBy").Value = ERPLoginData.User.ToUpper
            CLMasterSQLCommand1.Parameters("@SupplyType").Value = Header.SupplyType
            CLMasterSQLCommand1.Parameters("@ReasonCode").Value = Header.ReasonCode
            CLMasterSQLCommand1.Parameters("@OrgCode").Value = ERPLoginData.OrgCode
            'CLMasterSQLCommand1.Parameters("@MultiPOs").Value = 0

            CLMasterSQLCommand1.CommandType = CommandType.Text
            CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
            ra = CLMasterSQLCommand1.ExecuteNonQuery()

            For i = 0 To TOMydataSet.Tables("POItems").Rows.Count - 1
                CLMasterSQLCommand.Parameters("@PO").Value = TOMydataSet.Tables("POItems").Rows(i)("DJNo")
                CLMasterSQLCommand.Parameters("@Product").Value = TOMydataSet.Tables("POItems").Rows(i)("Product")
                CLMasterSQLCommand.Parameters("@BuildQty").Value = TOMydataSet.Tables("POItems").Rows(i)("IssueQty")
                CLMasterSQLCommand.CommandType = CommandType.Text
                ra = CLMasterSQLCommand.ExecuteNonQuery()
            Next

            Dim temp As String
            For j = 0 To TOMydataSet.Tables("ItemDetails").Rows.Count - 1
                CLMasterSQLCommand2.Parameters("@Item").Value = j + 1
                temp = TOMydataSet.Tables("ItemDetails").Rows(j)("Material")
                CLMasterSQLCommand2.Parameters("@Material").Value = temp.ToUpper
                CLMasterSQLCommand2.Parameters("@MaterialRevision").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Revision")
                CLMasterSQLCommand2.Parameters("@MPN").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("ManufacturerPN")
                CLMasterSQLCommand2.Parameters("@Manufacturer").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Manufacturer")
                CLMasterSQLCommand2.Parameters("@ReqQty").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("ReqQty")
                CLMasterSQLCommand2.Parameters("@PickedQty").Value = 0
                CLMasterSQLCommand2.Parameters("@Forming").Value = TOMydataSet.Tables("ItemDetails").Rows(j)("Forming")
                CLMasterSQLCommand2.CommandType = CommandType.Text
                CLMasterSQLCommand2.CommandTimeout = TimeOut_M5
                ra = CLMasterSQLCommand2.ExecuteNonQuery()
            Next
            myConn.Close()

            PostProdPicking.PDTO = NextID

        Catch ex As Exception
            ErrorLogging("MMC-PostProdPicking", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            PostProdPicking.ErrMessage = "Error:" & ex.Message & ex.Source
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try


    End Function

    Public Function DeleteTemplate(ByVal Template As String) As String
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CLMasterSQLCommand1 As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Try
            myConn.Open()

            CLMasterSQLCommand = New SqlClient.SqlCommand("DELETE FROM T_PDTOTemplateHeader where TOTemplate=@TOTemplate", myConn)
            CLMasterSQLCommand.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand.Parameters("@TOTemplate").Value = Template

            CLMasterSQLCommand1 = New SqlClient.SqlCommand("DELETE FROM T_PDTOTemplateItem where TOTemplate=@TOTemplate", myConn)
            CLMasterSQLCommand1.Parameters.Add("@TOTemplate", SqlDbType.NVarChar, 50, "TOTemplate")
            CLMasterSQLCommand1.Parameters("@TOTemplate").Value = Template

            CLMasterSQLCommand.CommandType = CommandType.Text
            ra = CLMasterSQLCommand.ExecuteNonQuery()
            CLMasterSQLCommand1.CommandType = CommandType.Text
            CLMasterSQLCommand1.CommandTimeout = TimeOut_M5
            ra = CLMasterSQLCommand1.ExecuteNonQuery()

            myConn.Close()
            DeleteTemplate = "okay"


        Catch ex As Exception
            ErrorLogging("MMC-DeleteTemplate", "", "Template: " & Template & ", " & ex.Message & ex.Source, "E")
            DeleteTemplate = "Error"
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function GetBOMFromERP(ByVal mydataset As DataSet, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim i As Integer
            Dim ds As New DataSet
            Dim BOMData, DJComponent As DataTable
            Dim myDataColumn As DataColumn
            Dim sda As SqlClient.SqlDataAdapter

            BOMData = New DataTable("BOMData")

            myDataColumn = New Data.DataColumn("Material", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("UOM", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)

            'myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.String"))
            'BOMData.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.String"))
            'BOMData.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.String"))
            'BOMData.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.String"))
            'BOMData.Columns.Add(myDataColumn)

            myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)

            myDataColumn = New Data.DataColumn("SubInv", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)

            myDataColumn = New Data.DataColumn("Forming", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)

            ds.Tables.Add(BOMData)

            'DJComponent = New DataTable("DJComponent")

            'myDataColumn = New Data.DataColumn("component", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("revision", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("component_desc", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("component_rev", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("required_quantity", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)

            'myDataColumn = New Data.DataColumn("quantity_issued", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("quantity_per_assembly", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("uom_code", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("sypply_type", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("supply_subinventory", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("supply_loc", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("qty_onhand", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("mpn", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("manufacturer", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("item_type", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("ejit_flag", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("kanban_flag", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("consumables_flag", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("bulky_flag", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("pre_work_flag", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)
            'myDataColumn = New Data.DataColumn("make_buy", System.Type.GetType("System.String"))
            'DJComponent.Columns.Add(myDataColumn)



            ds.Tables.Add("DJComponent")

            Dim myDataRow As Data.DataRow
            Dim qtydiff As Decimal
            Dim j As Integer
            Dim BuildQty As Integer = 0
            Dim n As Integer

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim Type_List As String
            Dim TypeList() As String
            Try


                'oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_lines"

                'oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                'oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar).Value = OrderNo
                'oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode

                'oda.SelectCommand.Connection.Open()
                'oda.Fill(ds, "DJComponent")
                'oda.SelectCommand.Connection.Close()

                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "dbo.ora_get_release_lines"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Parameters.Add("@p_discreate_job", SqlDbType.VarChar, 150).Value = OrderNo
                sda.SelectCommand.Parameters.Add("@p_org_id", SqlDbType.Int, 150).Value = ERPLoginData.OrgID
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds, "DJComponent")
                sda.SelectCommand.Connection.Close()

                Type_List = Get_TypeList(Header, ERPLoginData)

                'If Header.SupplyType.Trim.ToUpper = "OSP" OrElse Header.SupplyType.Trim.ToUpper = "FLOORSTOCK" Then
                '    For i = 0 To ds.Tables(1).Rows.Count - 1
                '        If FixNull(ds.Tables(1).Rows(i).Item("supply_subinventory")) <> "" AndAlso FixNull(ds.Tables(1).Rows(i).Item("supply_loc")) <> "" AndAlso FixNull(ds.Tables(1).Rows(i).Item("supply_loc")) <> "FGEK1...." Then
                '            Header.CheckSubInv = FixNull(ds.Tables(1).Rows(i).Item("supply_subinventory"))
                '            Header.CheckLocator = FixNull(ds.Tables(1).Rows(i).Item("supply_loc"))
                '            Exit For
                '        End If
                '    Next
                'ElseIf Header.SupplyType.Trim.ToUpper = "PULL" OrElse Header.SupplyType.Trim.ToUpper = "MSB" OrElse Header.SupplyType.Trim.ToUpper = "PRE-WORK" Then
                '    For i = 0 To ds.Tables(1).Rows.Count - 1
                '        If (FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "") Then
                '            Header.CheckSubInv = FixNull(ds.Tables(1).Rows(i).Item("supply_subinventory"))
                '            Header.CheckLocator = FixNull(ds.Tables(1).Rows(i).Item("supply_loc"))
                '            Exit For
                '        End If
                '    Next
                'End If

                For i = 0 To ds.Tables(1).Rows.Count - 1
                    If (Header.SupplyType.Trim.ToUpper = "PULL" OrElse Header.SupplyType.Trim.ToUpper = "MSB" OrElse Header.SupplyType.Trim.ToUpper = "PRE-WORK") AndAlso Not InStr(UCase(FixNull(ds.Tables(1).Rows(i).Item(8))), "PULL") > 0 Then
                        Continue For
                    ElseIf Header.SupplyType.Trim.ToUpper = "PUSH" AndAlso Not InStr(UCase(FixNull(ds.Tables(1).Rows(i).Item(8))), "PUSH") > 0 Then
                        Continue For
                    End If

                    If UCase(Type_List) = "NORMAL" AndAlso (FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "") Then

                    ElseIf InStr(UCase(Type_List), "MSB") > 0 AndAlso (FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "") AndAlso FixNull(ds.Tables(1).Rows(i).Item("manufacturer")) <> "" AndAlso FixNull(ds.Tables(1).Rows(i).Item("mpn")) <> "" AndAlso (((Header.MakeBuy = "Make" OrElse Header.MakeBuy = "All") AndAlso FixNull(ds.Tables(1).Rows(i).Item("make_buy")) = "Make") OrElse ((Header.MakeBuy = "Buy" OrElse Header.MakeBuy = "All") AndAlso FixNull(ds.Tables(1).Rows(i).Item("make_buy")) = "Buy")) Then

                    ElseIf InStr(UCase(Type_List), "PRE-WORK") > 0 AndAlso FixNull(ds.Tables(1).Rows(i).Item("pre_work_flag")) = "Y" AndAlso (FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "") Then

                    ElseIf InStr(UCase(Type_List), "EJIT") > 0 AndAlso FixNull(ds.Tables(1).Rows(i).Item("ejit_flag")) = "Y" Then

                    ElseIf InStr(UCase(Type_List), "KANBAN") > 0 AndAlso FixNull(ds.Tables(1).Rows(i).Item("kanban_flag")) = "Y" Then

                    ElseIf InStr(UCase(Type_List), "CONSUMABLE") > 0 AndAlso FixNull(ds.Tables(1).Rows(i).Item("consumables_flag")) = "Y" Then

                    ElseIf InStr(UCase(Type_List), "BULKY") > 0 AndAlso FixNull(ds.Tables(1).Rows(i).Item("bulky_flag")) = "Y" Then

                    ElseIf UCase(Type_List) = "ALL" Then

                    Else
                        Continue For
                    End If

                    myDataRow = BOMData.NewRow()
                    myDataRow("Material") = ds.Tables(1).Rows(i).Item("component")
                    If Not ds.Tables(1).Rows(i).Item("component_rev") Is DBNull.Value Then
                        myDataRow("Revision") = ds.Tables(1).Rows(i).Item("component_rev")
                    Else
                        myDataRow("Revision") = ""
                    End If
                    If Not ds.Tables(1).Rows(i).Item("uom_code") Is DBNull.Value Then
                        myDataRow("UOM") = ds.Tables(1).Rows(i).Item("uom_code")
                    Else
                        myDataRow("UOM") = ""
                    End If
                    If Not ds.Tables(1).Rows(i).Item("mpn") Is DBNull.Value Then
                        myDataRow("ManufacturerPN") = ds.Tables(1).Rows(i).Item("mpn")
                    Else
                        myDataRow("ManufacturerPN") = ""
                    End If
                    If Not ds.Tables(1).Rows(i).Item("manufacturer") Is DBNull.Value Then
                        myDataRow("Manufacturer") = ds.Tables(1).Rows(i).Item("manufacturer")
                    Else
                        myDataRow("Manufacturer") = ""
                    End If
                    If Not ds.Tables(1).Rows(i).Item("quantity_per_assembly") Is DBNull.Value Then
                        myDataRow("PerQty") = Convert.ToString(ds.Tables(1).Rows(i).Item("quantity_per_assembly"))
                    Else
                        myDataRow("PerQty") = 0
                    End If
                    'Change the way to get Available Qty, from eTrace, not Oracle.

                    'If Not ds.Tables(1).Rows(i).Item(11) Is DBNull.Value Then
                    '    myDataRow(6) = Convert.ToDouble(ds.Tables(1).Rows(i).Item(11))
                    'Else
                    '    myDataRow(6) = 0
                    'End If
                    myDataRow("AvlQty") = Get_AvlQty(myDataRow("Material"), myDataRow("Revision"), myDataRow("Manufacturer"), myDataRow("ManufacturerPN"), Header, ERPLoginData)

                    BuildQty = 0
                    For j = 0 To mydataset.Tables(0).Rows.Count - 1
                        BuildQty = BuildQty + mydataset.Tables(0).Rows(j)("IssueQty")
                    Next
                    If mydataset.Tables(0).Rows.Count > 1 Then
                        For n = 0 To mydataset.Tables(0).Rows.Count - 2
                            If mydataset.Tables(0).Rows(n)(1) = mydataset.Tables(0).Rows(n + 1)(1) Then
                                myDataRow("ReqQty") = myDataRow("PerQty") * BuildQty
                            Else
                                myDataRow("ReqQty") = myDataRow("PerQty") * Header.BuildQty
                            End If
                        Next
                    Else
                        myDataRow("ReqQty") = myDataRow("PerQty") * BuildQty
                    End If

                    qtydiff = myDataRow("ReqQty") - myDataRow("AvlQty")
                    myDataRow("Shortage") = IIf(qtydiff > 0, qtydiff, 0)

                    If Header.SupplyType.Trim.ToUpper = "OSP" OrElse Header.SupplyType.Trim.ToUpper = "FLOORSTOCK" Then
                        myDataRow("SubInv") = ds.Tables(1).Rows(i).Item("supply_subinventory")
                        myDataRow("Locator") = ds.Tables(1).Rows(i).Item("supply_loc")
                    End If

                    myDataRow("Forming") = ds.Tables(1).Rows(i).Item("pre_work_flag")

                    BOMData.Rows.Add(myDataRow)
                Next

            Catch ex As Exception
                ErrorLogging("MMC-GetBOMFromERP", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
            Return ds
        End Using

    End Function

    Public Function GetListFromExcel(ByVal mydataset As DataSet, ByVal ExcelData As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As PPDataRst
        Using da As DataAccess = GetDataAccess()
            Dim i As Integer
            Dim ds As New DataSet
            Dim BOMData As DataTable
            Dim myDataColumn As DataColumn
            Dim myDataRow As Data.DataRow
            Dim qtydiff As Decimal
            Dim j As Integer
            Dim BuildQty As Integer = 0
            Dim n As Integer
            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim DR1() As DataRow = Nothing
            Dim DR2() As DataRow = Nothing
            Dim DR3() As DataRow = Nothing
            Dim DR4() As DataRow = Nothing
            Dim ErrMsg As String

            GetListFromExcel.dsList = New DataSet
            GetListFromExcel.ErrMsg = ""

            BOMData = New DataTable("BOMData")
            myDataColumn = New Data.DataColumn("Material", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.Decimal"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("SubInv", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            ds.Tables.Add(BOMData)

            Try
                DR1 = ExcelData.Tables(0).Select("(item is null or item = '') and (qtyper is null or qtyper = '')")
                If DR1.Length > 0 Then
                    For Each drblank As DataRow In DR1
                        ExcelData.Tables(0).Rows.Remove(drblank)
                    Next
                End If

                If Header.SupplyType.ToUpper = "MSB" Then
                    DR2 = ExcelData.Tables(0).Select("(manufacturer is null or manufacturer = '') or (manufacturerpn is null or manufacturerpn = '')")
                    If DR2.Length > 0 Then
                        ErrMsg = "Manufacturer & ManufacturerPN must be entered when SupplyType = 'MSB'."
                    End If
                End If

                DR3 = ExcelData.Tables(0).Select("item is null or item = ''")
                If DR3.Length > 0 Then
                    ErrMsg = ErrMsg & "Item must be entered"
                End If

                DR4 = ExcelData.Tables(0).Select("qtyper = 0")
                If DR4.Length > 0 Then
                    ErrMsg = ErrMsg & "QtyPer must be entered"
                End If

                If ErrMsg <> "" Then
                    GetListFromExcel.ErrMsg = ErrMsg
                    Exit Function
                End If

                For i = 0 To ExcelData.Tables(0).Rows.Count - 1
                    myDataRow = BOMData.NewRow()
                    myDataRow("Material") = UCase(IIf(ExcelData.Tables(0).Rows(i)(0) Is DBNull.Value, "", ExcelData.Tables(0).Rows(i)(0).ToString.Trim))
                    myDataRow("Revision") = UCase(IIf(ExcelData.Tables(0).Rows(i)(1) Is DBNull.Value, "", ExcelData.Tables(0).Rows(i)(1).ToString.Trim))
                    myDataRow("Manufacturer") = IIf(ExcelData.Tables(0).Rows(i)(2) Is DBNull.Value, "", ExcelData.Tables(0).Rows(i)(2).ToString.Trim)
                    myDataRow("ManufacturerPN") = IIf(ExcelData.Tables(0).Rows(i)(3) Is DBNull.Value, "", ExcelData.Tables(0).Rows(i)(3).ToString.Trim)
                    If FixNull(ExcelData.Tables(0).Rows(i)(4)) = "" OrElse FixNull(ExcelData.Tables(0).Rows(i)(4)) = 0 Then
                        'myDataRow("PerQty") = 0
                        ErrMsg = ErrMsg & "Column 'QtyPer' must be entered"
                        GetListFromExcel.ErrMsg = ErrMsg
                        GetListFromExcel.dsList.Clear()
                        Exit Function
                    Else
                        myDataRow("PerQty") = CDec(Format(ExcelData.Tables(0).Rows(i)(4), "#.######"))
                    End If
                    'Change the way to get Available Qty, from eTrace, not Oracle.

                    'If Not ds.Tables(1).Rows(i).Item(11) Is DBNull.Value Then
                    '    myDataRow(6) = Convert.ToDouble(ds.Tables(1).Rows(i).Item(11))
                    'Else
                    '    myDataRow(6) = 0
                    'End If
                    myDataRow("AvlQty") = Get_AvlQty(myDataRow("Material"), myDataRow("Revision"), myDataRow("Manufacturer"), myDataRow("ManufacturerPN"), Header, ERPLoginData)

                    BuildQty = 0
                    For j = 0 To mydataset.Tables(0).Rows.Count - 1
                        BuildQty = BuildQty + mydataset.Tables(0).Rows(j)("IssueQty")
                    Next
                    If mydataset.Tables(0).Rows.Count > 1 Then
                        For n = 0 To mydataset.Tables(0).Rows.Count - 2
                            If mydataset.Tables(0).Rows(n)(1) = mydataset.Tables(0).Rows(n + 1)(1) Then
                                myDataRow("ReqQty") = myDataRow("PerQty") * BuildQty
                            Else
                                myDataRow("ReqQty") = myDataRow("PerQty") * Header.BuildQty
                            End If
                        Next
                    Else
                        myDataRow("ReqQty") = myDataRow("PerQty") * BuildQty
                    End If

                    qtydiff = myDataRow("ReqQty") - myDataRow("AvlQty")
                    myDataRow("Shortage") = IIf(qtydiff > 0, qtydiff, 0)

                    'If Header.SupplyType.Trim.ToUpper = "OSP" OrElse Header.SupplyType.Trim.ToUpper = "FLOORSTOCK" Then
                    '    myDataRow("SubInv") = ds.Tables(1).Rows(i).Item("supply_subinventory")
                    '    myDataRow("Locator") = ds.Tables(1).Rows(i).Item("supply_loc")
                    'End If
                    BOMData.Rows.Add(myDataRow)
                Next
                GetListFromExcel.ErrMsg = ErrMsg
                GetListFromExcel.dsList = ds.Copy

            Catch ex As Exception
                GetListFromExcel.ErrMsg = ex.Message.ToString
                GetListFromExcel.dsList.Clear()
                ErrorLogging("MMC-GetListFromExcel", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_DestSubLoc(ByVal DJ As String, ByVal ERPLoginData As ERPLogin, ByVal Header As ProdPickingStructure) As ProdPickingStructure
        Using da As DataAccess = GetDataAccess()

            Dim i As Integer
            Dim ds, Dest As New DataSet

            Dim BuildQty As Integer = 0
            Dim n As Integer

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim sda As SqlClient.SqlDataAdapter
            Try
                'oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_lines"

                'oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                'oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar).Value = DJ
                'oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode

                'oda.SelectCommand.Connection.Open()
                'oda.Fill(ds)
                'oda.SelectCommand.Connection.Close()

                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "dbo.ora_get_release_lines"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Parameters.Add("@p_discreate_job", SqlDbType.VarChar, 150).Value = DJ
                sda.SelectCommand.Parameters.Add("@p_org_id", SqlDbType.Int, 150).Value = ERPLoginData.OrgID
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds)
                sda.SelectCommand.Connection.Close()

                If Header.SupplyType.Trim.ToUpper = "OSP" OrElse Header.SupplyType.Trim.ToUpper = "FLOORSTOCK" Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If FixNull(ds.Tables(0).Rows(i).Item("supply_subinventory")) <> "" AndAlso FixNull(ds.Tables(0).Rows(i).Item("supply_loc")) <> "" Then
                            Header.CheckSubInv = FixNull(ds.Tables(0).Rows(i).Item("supply_subinventory"))
                            Header.CheckLocator = FixNull(ds.Tables(0).Rows(i).Item("supply_loc"))
                            Exit For
                        End If
                    Next
                    If Header.SupplyType.Trim.ToUpper = "FLOORSTOCK" Then
                        Dest = GetOrderInfoFromERP(ERPLoginData, DJ)
                        Header.CheckLocator = Dest.Tables(0).Rows(0)("prod_line_desc").ToString.Trim
                    End If
                ElseIf Header.SupplyType.Trim.ToUpper = "PULL" OrElse Header.SupplyType.Trim.ToUpper = "MSB" OrElse Header.SupplyType.Trim.ToUpper = "PRE-WORK" Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If FixNull(ds.Tables(0).Rows(i).Item("supply_subinventory")) <> "" AndAlso (FixNull(ds.Tables(0).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("bulky_flag")) = "") Then
                            Header.CheckSubInv = FixNull(ds.Tables(0).Rows(i).Item("supply_subinventory"))
                            Header.CheckLocator = FixNull(ds.Tables(0).Rows(i).Item("supply_loc"))
                            Exit For
                        End If
                    Next
                End If
                Return Header
            Catch ex As Exception
                ErrorLogging("MMC-GetBOMFromERP", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_ChangeExcelFlag(ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK011"))

                Get_ChangeExcelFlag = FixNull(result.ToString)
            Catch ex As Exception
                ErrorLogging("MMC-Get_ChangeExcelFlag", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function Get_TypeList(ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                If UCase(Header.SupplyType) = "PULL" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK004"))
                ElseIf UCase(Header.SupplyType) = "PUSH" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK005"))
                ElseIf UCase(Header.SupplyType) = "OSP" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK006"))
                ElseIf UCase(Header.SupplyType) = "MSB" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK007"))
                ElseIf UCase(Header.SupplyType) = "FLOORSTOCK" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK003"))
                ElseIf UCase(Header.SupplyType) = "PRE-WORK" Then
                    result = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = '{0}' ", "PIK008"))
                End If
                Get_TypeList = FixNull(result.ToString)
            Catch ex As Exception
                ErrorLogging("MMC-Get_TypeList", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function Get_SrcSubInv(ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                Dim SqlStr, SubInvList As String

                SqlStr = String.Format("SELECT value from T_Config where configID = 'PIK010'")
                SubInvList = FixNull(da.ExecuteScalar(SqlStr))
                Return SubInvList

            Catch ex As Exception
                ErrorLogging("MMC-Get_SrcSubInv", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function Get_AvlQty(ByVal Material As String, ByVal Rev As String, ByVal Manufacturer As String, ByVal ManufacturerPN As String, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As Decimal
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                Dim SqlStr, SubInvList As String

                SqlStr = String.Format("SELECT value from T_Config where configID = 'PIK010'")
                SubInvList = FixNull(da.ExecuteScalar(SqlStr))

                If Rev <> "" Then
                    If UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList = "" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and MaterialRevision = '{4}' and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%' and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode, Rev))
                    ElseIf UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList <> "" AndAlso Header.SrcSubInv = "All" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and MaterialRevision = '{4}' and StatusCode = 1 and (CHARINDEX(SLOC, (SELECT Value FROM T_Config WHERE (ConfigID = 'PIK010'))) > 0) and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode, Rev))
                    ElseIf UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList <> "" AndAlso Header.SrcSubInv <> "All" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and MaterialRevision = '{4}' and StatusCode = 1 and SLOC = '{5}' and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode, Rev, Header.SrcSubInv))
                    ElseIf UCase(Header.SupplyType) = "PULL" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%'", Material, ERPLoginData.OrgCode, Rev))
                    ElseIf UCase(Header.SupplyType) = "PRE-WORK" Then
                        Dim SubInv, Locator As String
                        SubInv = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT004'"))
                        Locator = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT005'"))
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC = '{3}' and StorageBin = '{4}'", Material, ERPLoginData.OrgCode, Rev, SubInv, Locator))
                    ElseIf UCase(Header.SupplyType) = "OSP" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC = '{3}' and StorageBin = '{4}'", Material, ERPLoginData.OrgCode, Rev, Header.CheckSubInv, Header.CheckLocator))
                    ElseIf UCase(Header.SupplyType) = "PUSH" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%'", Material, ERPLoginData.OrgCode, Rev))
                    ElseIf UCase(Header.SupplyType) = "FLOORSTOCK" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC = '{3}' and StorageBin = '{4}'", Material, ERPLoginData.OrgCode, Rev, Header.CheckSubInv, Header.CheckLocator))
                    End If
                Else
                    If UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList = "" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%' and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode))
                    ElseIf UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList <> "" AndAlso Header.SrcSubInv = "All" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and (CHARINDEX(SLOC, (SELECT Value FROM T_Config WHERE (ConfigID = 'PIK010'))) > 0) and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode))
                    ElseIf UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" AndAlso SubInvList <> "" AndAlso Header.SrcSubInv <> "All" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{3}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC = '{4}' and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode, Header.SrcSubInv))
                    ElseIf UCase(Header.SupplyType) = "PULL" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%'", Material, ERPLoginData.OrgCode))
                    ElseIf UCase(Header.SupplyType) = "PRE-WORK" Then
                        Dim SubInv, Locator As String
                        SubInv = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT004'"))
                        Locator = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT005'"))
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC = '{2}' and StorageBin = '{3}'", Material, ERPLoginData.OrgCode, SubInv, Locator))
                    ElseIf UCase(Header.SupplyType) = "OSP" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC = '{2}' and StorageBin = '{3}'", Material, ERPLoginData.OrgCode, Header.CheckSubInv, Header.CheckLocator))
                    ElseIf UCase(Header.SupplyType) = "PUSH" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%'", Material, ERPLoginData.OrgCode))
                    ElseIf UCase(Header.SupplyType) = "FLOORSTOCK" Then
                        result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where OrgCode = '{1}' and MaterialNo = '{0}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC = '{2}' and StorageBin = '{3}'", Material, ERPLoginData.OrgCode, Header.CheckSubInv, Header.CheckLocator))
                    End If
                End If
                'If UCase(Header.SupplyType) = "MSB" AndAlso Manufacturer <> "" AndAlso ManufacturerPN <> "" Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{3}' and MaterialNo = '{0}' and MaterialRevision = '{4}' and StatusCode = 1 and SLOC not like '%OS%' and Manufacturer = '{1}' and ManufacturerPN = '{2}' ", Material, Manufacturer, ManufacturerPN, ERPLoginData.OrgCode, Rev))
                'ElseIf UCase(Header.SupplyType) = "MSB" AndAlso (Manufacturer = "" OrElse ManufacturerPN = "") Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%'", Material, ERPLoginData.OrgCode, Rev))
                'ElseIf UCase(Header.SupplyType) = "PULL" Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%'", Material, ERPLoginData.OrgCode, Rev))
                'ElseIf UCase(Header.SupplyType) = "OSP" Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC like '%OS%'", Material, ERPLoginData.OrgCode, Rev))
                'ElseIf UCase(Header.SupplyType) = "PUSH" Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%'", Material, ERPLoginData.OrgCode, Rev))
                'ElseIf UCase(Header.SupplyType) = "FLOORSTOCK" Then
                '    result = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster where OrgCode = '{1}' and MaterialNo = '{0}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC like '%FS%'", Material, ERPLoginData.OrgCode, Rev))
                'End If

                If Not result Is Nothing AndAlso result.ToString <> "" Then
                    Get_AvlQty = result.ToString
                Else
                    Get_AvlQty = 0
                End If
            Catch ex As Exception
                ErrorLogging("MMC-Get_AvlQty", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetAvlQty(ByVal Material As String, ByVal ERPLoginData As ERPLogin) As Get_Qty
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result1 As Object
                result1 = da.ExecuteScalar(String.Format("select Sum(QtyBaseUOM) from T_CLMaster with (nolock) where StatusCode = 1 and MaterialNo = '{0}' and StorageBin NOT LIKE 'TR%'", Material))
                If result1.ToString <> "" Then
                    'If Not result1.ToString Is DBNull.Value Then
                    GetAvlQty.AvlQty = result1.ToString
                Else
                    GetAvlQty.AvlQty = 0
                End If
                'result2 = da.ExecuteScalar(String.Format("Select Sum(QtyBaseUOM) from T_CLMaster where Sloc = '0010' and StatusCode = 1 and StorageType = '001' and MaterialNo = '{0}' and StorageBin NOT LIKE 'TR%'", Material))
                'GetAvlQty.WHSQty = result2.ToString
            Catch ex As Exception
                ErrorLogging("MMC-GetAvlQty", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetMSBSource(ByVal mydataset As DataSet, ByVal DJItems As DataSet, ByVal Header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim dsMSB As New DataSet
            Dim ds As DataSet
            Dim MSBSource As DataTable
            Dim myDataColumn As DataColumn
            Dim myDataRow As DataRow
            Dim i, j As Integer

            Try
                MSBSource = New DataTable("MSBSource")
                myDataColumn = New Data.DataColumn("Material", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("UOM", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("PickedQty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("OpenQty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Status", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("DJNo", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("DJRevision", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("DJQty", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Model", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("ProdFloor", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("SubInv", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Locator", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("DestSubinv", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("RTLot", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("ExpDate", System.Type.GetType("System.String"))
                MSBSource.Columns.Add(myDataColumn)
                myDataColumn = New Data.DataColumn("Qty", System.Type.GetType("System.Decimal"))
                MSBSource.Columns.Add(myDataColumn)
                dsMSB.Tables.Add(MSBSource)

                Dim SqlStr, SubInvList As String
                Dim DR() As DataRow = Nothing
                SqlStr = String.Format("SELECT value from T_Config where configID = 'PIK010'")
                SubInvList = FixNull(da.ExecuteScalar(SqlStr))

                If SubInvList <> "" AndAlso Header.SrcSubInv = "All" Then
                    For i = 0 To mydataset.Tables("ItemDetails").Rows.Count - 1
                        If mydataset.Tables("ItemDetails").Rows(i)("Revision") = "" Then
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and (CHARINDEX(SLOC, (SELECT Value FROM T_Config WHERE (ConfigID = 'PIK010'))) > 0) and Manufacturer = '{2}' and ManufacturerPN = '{3}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                        Else
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and MaterialRevision = '{2}' and StatusCode = 1 and (CHARINDEX(SLOC, (SELECT Value FROM T_Config WHERE (ConfigID = 'PIK010'))) > 0) and Manufacturer = '{3}' and ManufacturerPN = '{4}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), mydataset.Tables("ItemDetails").Rows(i)("Revision"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                        End If
                        ds = New DataSet
                        ds = da.ExecuteDataSet(SqlStr, "SLOC")
                        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                            For j = 0 To ds.Tables(0).Rows.Count - 1
                                myDataRow = dsMSB.Tables("MSBSource").NewRow
                                myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                                myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                                DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                                If DR.Length > 0 Then
                                    myDataRow("UOM") = FixNull(DR(0)("uom"))
                                Else
                                    myDataRow("UOM") = ""
                                End If
                                'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                                myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                                myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                                myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                                myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                                myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                                myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                                myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                                myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                                myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                                myDataRow("DJNo") = Header.PO
                                myDataRow("DJRevision") = Header.DJRev
                                myDataRow("DJQty") = Header.DJQty
                                myDataRow("Model") = Header.Product
                                myDataRow("ProdFloor") = Header.ProdFloor
                                myDataRow("DestSubinv") = Header.CheckSubInv

                                myDataRow("SubInv") = FixNull(ds.Tables(0).Rows(j)("SLOC"))
                                myDataRow("Locator") = FixNull(ds.Tables(0).Rows(j)("StorageBin"))
                                myDataRow("RTLot") = FixNull(ds.Tables(0).Rows(j)("RTLot"))
                                myDataRow("ExpDate") = FixNull(ds.Tables(0).Rows(j)("ExpDate"))
                                myDataRow("Qty") = FixNull(ds.Tables(0).Rows(j)("Qty"))

                                dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                            Next
                        Else
                            myDataRow = dsMSB.Tables("MSBSource").NewRow
                            myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                            myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                            DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                            If DR.Length > 0 Then
                                myDataRow("UOM") = FixNull(DR(0)("uom"))
                            Else
                                myDataRow("UOM") = ""
                            End If
                            'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                            myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                            myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                            myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                            myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                            myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                            myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                            myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                            myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                            myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                            myDataRow("DJNo") = Header.PO
                            myDataRow("DJRevision") = Header.DJRev
                            myDataRow("DJQty") = Header.DJQty
                            myDataRow("Model") = Header.Product
                            myDataRow("ProdFloor") = Header.ProdFloor
                            myDataRow("DestSubinv") = Header.CheckSubInv

                            myDataRow("SubInv") = ""
                            myDataRow("Locator") = ""
                            myDataRow("RTLot") = ""
                            myDataRow("ExpDate") = ""
                            myDataRow("Qty") = 0

                            dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                        End If

                    Next
                ElseIf SubInvList <> "" AndAlso Header.SrcSubInv <> "All" Then
                    For i = 0 To mydataset.Tables("ItemDetails").Rows.Count - 1
                        If mydataset.Tables("ItemDetails").Rows(i)("Revision") = "" Then
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC = '{4}' and Manufacturer = '{2}' and ManufacturerPN = '{3}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"), Header.SrcSubInv)
                        Else
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC = '{5}' and Manufacturer = '{3}' and ManufacturerPN = '{4}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), mydataset.Tables("ItemDetails").Rows(i)("Revision"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"), Header.SrcSubInv)
                        End If
                        ds = New DataSet
                        ds = da.ExecuteDataSet(SqlStr, "SLOC")
                        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                            For j = 0 To ds.Tables(0).Rows.Count - 1
                                myDataRow = dsMSB.Tables("MSBSource").NewRow
                                myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                                myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                                DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                                If DR.Length > 0 Then
                                    myDataRow("UOM") = FixNull(DR(0)("uom"))
                                Else
                                    myDataRow("UOM") = ""
                                End If
                                'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                                myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                                myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                                myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                                myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                                myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                                myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                                myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                                myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                                myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                                myDataRow("DJNo") = Header.PO
                                myDataRow("DJRevision") = Header.DJRev
                                myDataRow("DJQty") = Header.DJQty
                                myDataRow("Model") = Header.Product
                                myDataRow("ProdFloor") = Header.ProdFloor
                                myDataRow("DestSubinv") = Header.CheckSubInv

                                myDataRow("SubInv") = FixNull(ds.Tables(0).Rows(j)("SLOC"))
                                myDataRow("Locator") = FixNull(ds.Tables(0).Rows(j)("StorageBin"))
                                myDataRow("RTLot") = FixNull(ds.Tables(0).Rows(j)("RTLot"))
                                myDataRow("ExpDate") = FixNull(ds.Tables(0).Rows(j)("ExpDate"))
                                myDataRow("Qty") = FixNull(ds.Tables(0).Rows(j)("Qty"))

                                dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                            Next
                        Else
                            myDataRow = dsMSB.Tables("MSBSource").NewRow
                            myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                            myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                            DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                            If DR.Length > 0 Then
                                myDataRow("UOM") = FixNull(DR(0)("uom"))
                            Else
                                myDataRow("UOM") = ""
                            End If
                            'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                            myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                            myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                            myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                            myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                            myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                            myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                            myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                            myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                            myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                            myDataRow("DJNo") = Header.PO
                            myDataRow("DJRevision") = Header.DJRev
                            myDataRow("DJQty") = Header.DJQty
                            myDataRow("Model") = Header.Product
                            myDataRow("ProdFloor") = Header.ProdFloor
                            myDataRow("DestSubinv") = Header.CheckSubInv

                            myDataRow("SubInv") = ""
                            myDataRow("Locator") = ""
                            myDataRow("RTLot") = ""
                            myDataRow("ExpDate") = ""
                            myDataRow("Qty") = 0

                            dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                        End If

                    Next
                ElseIf SubInvList = "" Then
                    For i = 0 To mydataset.Tables("ItemDetails").Rows.Count - 1
                        If mydataset.Tables("ItemDetails").Rows(i)("Revision") = "" Then
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and (MaterialRevision = '' or MaterialRevision is null) and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%' and Manufacturer = '{2}' and ManufacturerPN = '{3}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                        Else
                            SqlStr = String.Format("select SLOC, StorageBin, RTLot, ExpDate, Sum(QtyBaseUOM) as Qty from T_CLMaster where OrgCode = '{0}' and MaterialNo = '{1}' and MaterialRevision = '{2}' and StatusCode = 1 and SLOC not like '%OS%' and SLOC not like '%BF%' and SLOC not like '%SCR%' and SLOC not like '%MRB%' and SLOC not like '%WO%' and SLOC not like '%RTV%' and Manufacturer = '{3}' and ManufacturerPN = '{4}' group by SLOC, StorageBin, RTLot, ExpDate", ERPLoginData.OrgCode, mydataset.Tables("ItemDetails").Rows(i)("Material"), myDataRow("Revision"), mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"), mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                        End If
                        ds = New DataSet
                        ds = da.ExecuteDataSet(SqlStr, "SLOC")
                        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                            For j = 0 To ds.Tables(0).Rows.Count - 1
                                myDataRow = dsMSB.Tables("MSBSource").NewRow
                                myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                                myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                                DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                                If DR.Length > 0 Then
                                    myDataRow("UOM") = FixNull(DR(0)("uom"))
                                Else
                                    myDataRow("UOM") = ""
                                End If
                                'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                                myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                                myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                                myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                                myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                                myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                                myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                                myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                                myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                                myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                                myDataRow("DJNo") = Header.PO
                                myDataRow("DJRevision") = Header.DJRev
                                myDataRow("DJQty") = Header.DJQty
                                myDataRow("Model") = Header.Product
                                myDataRow("ProdFloor") = Header.ProdFloor
                                myDataRow("DestSubinv") = Header.CheckSubInv

                                myDataRow("SubInv") = FixNull(ds.Tables(0).Rows(j)("SLOC"))
                                myDataRow("Locator") = FixNull(ds.Tables(0).Rows(j)("StorageBin"))
                                myDataRow("RTLot") = FixNull(ds.Tables(0).Rows(j)("RTLot"))
                                myDataRow("ExpDate") = FixNull(ds.Tables(0).Rows(j)("ExpDate"))
                                myDataRow("Qty") = FixNull(ds.Tables(0).Rows(j)("Qty"))

                                dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                            Next
                        Else
                            myDataRow = dsMSB.Tables("MSBSource").NewRow
                            myDataRow("Material") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material"))
                            myDataRow("Revision") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Revision"))

                            DR = DJItems.Tables("BOMData").Select(" material = '" & FixNull(mydataset.Tables("ItemDetails").Rows(i)("Material")) & "'")
                            If DR.Length > 0 Then
                                myDataRow("UOM") = FixNull(DR(0)("uom"))
                            Else
                                myDataRow("UOM") = ""
                            End If
                            'myDataRow("UOM") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("UOM"))

                            myDataRow("ManufacturerPN") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("ManufacturerPN"))
                            myDataRow("Manufacturer") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Manufacturer"))
                            myDataRow("PerQty") = mydataset.Tables("ItemDetails").Rows(i)("PerQty")
                            myDataRow("ReqQty") = mydataset.Tables("ItemDetails").Rows(i)("ReqQty")
                            myDataRow("AvlQty") = mydataset.Tables("ItemDetails").Rows(i)("AvlQty")
                            myDataRow("PickedQty") = mydataset.Tables("ItemDetails").Rows(i)("PickedQty")
                            myDataRow("OpenQty") = mydataset.Tables("ItemDetails").Rows(i)("OpenQty")
                            myDataRow("Shortage") = mydataset.Tables("ItemDetails").Rows(i)("Shortage")
                            myDataRow("Status") = FixNull(mydataset.Tables("ItemDetails").Rows(i)("Status"))

                            myDataRow("DJNo") = Header.PO
                            myDataRow("DJRevision") = Header.DJRev
                            myDataRow("DJQty") = Header.DJQty
                            myDataRow("Model") = Header.Product
                            myDataRow("ProdFloor") = Header.ProdFloor
                            myDataRow("DestSubinv") = Header.CheckSubInv

                            myDataRow("SubInv") = ""
                            myDataRow("Locator") = ""
                            myDataRow("RTLot") = ""
                            myDataRow("ExpDate") = ""
                            myDataRow("Qty") = 0

                            dsMSB.Tables("MSBSource").Rows.Add(myDataRow)
                        End If
                    Next
                End If
                Return dsMSB
            Catch ex As Exception
                ErrorLogging("MMC-GetMSBSource", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetOrderInfoFromERP(ByVal ERPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("DJInfo")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_dj"

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = OrderNo.Trim
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "DJInfo")
                oda.SelectCommand.Connection.Close()
                Return ds
            Catch oe As Exception
                Throw oe
                ErrorLogging("Material replenishment-GetOrderInfoFromERP", "", "DJ: " & OrderNo & ", " & oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetDJInfoFromERP(ByVal ERPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("DJInfo")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_dj_header"

                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar, 1000).Value = OrderNo.Trim
                oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "DJInfo")
                oda.SelectCommand.Connection.Close()
                Return ds
            Catch oe As Exception
                Throw oe
                ErrorLogging("Material replenishment-GetDJInfoFromERP", "", "DJ: " & OrderNo & ", " & oe.Message & oe.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function GetOrderInfoFromETRACE(ByVal ERPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sqlstr As String
                sqlstr = String.Format("SELECT top 1  PO, Model as 'PRODUCT_NUMBER', PCBA, POQty as 'OPEN_QUANTITY', POQty as 'ProdQty' FROM T_DJInfo with (nolock) WHERE PO = '{0}' and OrgCode = '{1}'", OrderNo, ERPLoginData.OrgCode)
                Return da.ExecuteDataSet(sqlstr, "DJInfo")
            Catch ex As Exception
                Throw ex
                ErrorLogging("Material replenishment-GetOrderInfoFromETRACE", ERPLoginData.User, "DJ: " & OrderNo & ", " & ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetServerFlag(ByVal ERPLoginData As ERPLogin) As String
        Dim Flag As String
        Using da As DataAccess = GetDataAccess()
            Try
                Flag = da.ExecuteScalar(String.Format("select value from t_config with (nolock) where ConfigID = 'MMT012'"))
                Return Flag
            Catch ex As Exception
                ErrorLogging("MMC-GetServerFlag", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Return ""
            End Try
        End Using
    End Function

    Public Function GetIDInfo(ByVal ERPLoginData As ERPLogin, ByVal ID As String) As DataSet
        GetIDInfo = New DataSet
        Dim DJNO, Flag As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result1 As New Object
                Dim result2 As New Object

                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                    'ID.Contains("B") Then
                    'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                    result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1", ID))

                    result2 = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where CLID = '{0}' and statuscode = 1", ID))
                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    'ID.Contains("P") Then
                    'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                    result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1", ID))

                    result2 = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where BoxID = '{0}' and statuscode = 1", ID))
                End If

                'Check record in APP68
                Flag = FixNull(GetServerFlag(ERPLoginData))
                If Flag = "YES" Then
                    If result1 Is Nothing Then
                        If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                            'ID.Contains("B") Then
                            'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                            result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1", ID))

                        ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                            'ID.Contains("P") Then
                            result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1", ID))
                        End If
                    End If
                End If

                If result2 Is Nothing Then
                    'If result1.ToString <> "" Then
                    '2010-9-8 Yong
                    If Not result1 Is Nothing Then
                        DJNO = result1.ToString
                    Else
                        DJNO = ""
                        Exit Function
                    End If
                    GetIDInfo = GetOrderInfoFromERP(ERPLoginData, DJNO)
                Else
                    Exit Function
                End If

            Catch ex As Exception
                ErrorLogging("MMC-GetIDInfo", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetDJ(ByVal ERPLoginData As ERPLogin, ByVal ID As String) As String

        Using da As DataAccess = GetDataAccess()
            Try
                Dim result1 As New Object
                Dim result2 As New Object
                Dim Flag As String

                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                    'ID.Contains("B") Then
                    'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                    result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1", ID))

                    result2 = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where CLID = '{0}' and statuscode = 1", ID))
                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    'ID.Contains("P") Then
                    'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                    result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1", ID))

                    result2 = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where BoxID = '{0}' and statuscode = 1", ID))
                End If

                'Check record in APP68
                Flag = FixNull(GetServerFlag(ERPLoginData))
                If Flag = "YES" Then
                    If result1 Is Nothing Then
                        If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                            'ID.Contains("B") Then
                            'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1 and (OrgCode is NULL or OrgCode = '{1}')", ID, ERPLoginData.OrgCode))
                            result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and statuscode = 1", ID))

                        ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                            'ID.Contains("P") Then
                            result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and statuscode = 1", ID))
                        End If
                    End If
                End If

                If result2 Is Nothing Then
                    'If result1.ToString <> "" Then
                    'If Not result1.ToString Is DBNull.Value Then
                    '2010-9-8 Yong
                    If Not result1 Is Nothing Then
                        GetDJ = result1.ToString
                    Else
                        GetDJ = ""
                    End If
                Else
                    Exit Function
                End If

            Catch ex As Exception
                ErrorLogging("MMC-GetDJ", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function Check_HW_Flag(ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                result = da.ExecuteScalar(String.Format("select value from T_Config where configid = '{0}'", "CLID014"))
                Check_HW_Flag = FixNull(result.ToString)
            Catch ex As Exception
                ErrorLogging("MMC-Check_HW_Flag", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Check_HW_Flag = ""
            End Try
        End Using
    End Function

    Public Function Flag_AutoGetBOMRev(ByVal ERPLoginData As ERPLogin) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim result As Object
                result = da.ExecuteScalar(String.Format("select value from T_Config where configid = '{0}'", "CLID008"))
                Flag_AutoGetBOMRev = FixNull(result.ToString)
            Catch ex As Exception
                ErrorLogging("MMC-Flag_AutoGetBOMRev", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Flag_AutoGetBOMRev = ""
            End Try
        End Using
    End Function

    Public Function CountOfBoxID(ByVal ERPLoginData As ERPLogin, ByVal ID As String) As DataSet
        Dim dsCount As DataSet
        Using da As DataAccess = GetDataAccess()
            Try
                Dim Sqlstr As String = ""
                Dim Flag As String
                '2010-8-10 Yong
                'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                '    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from T_Shippment with (nolock) where CartonID = '{0}' and (OrgCode is NULL or OrgCode = '{1}') and StatusCode = '1' GROUP BY PalletID, CartonID ", ID, ERPLoginData.OrgCode))
                'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                '    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from T_Shippment with (nolock) where PalletID = '{0}' and (OrgCode is NULL or OrgCode = '{1}') and StatusCode = '1' GROUP BY PalletID, CartonID ", ID, ERPLoginData.OrgCode))
                'End If

                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' GROUP BY PalletID, CartonID ", ID))
                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' GROUP BY PalletID, CartonID ", ID))
                End If


                dsCount = New DataSet
                dsCount = da.ExecuteDataSet(Sqlstr, "IDInfor")

                'check record in APP68
                Flag = FixNull(GetServerFlag(ERPLoginData))
                If Flag = "YES" Then
                    If dsCount Is Nothing OrElse dsCount.Tables.Count < 1 OrElse dsCount.Tables(0).Rows.Count < 1 Then
                        'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                        '    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and (OrgCode is NULL or OrgCode = '{1}') and StatusCode = '1' GROUP BY PalletID, CartonID ", ID, ERPLoginData.OrgCode))
                        'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                        '    Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and (OrgCode is NULL or OrgCode = '{1}') and StatusCode = '1' GROUP BY PalletID, CartonID ", ID, ERPLoginData.OrgCode))
                        'End If

                        If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                            Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' GROUP BY PalletID, CartonID ", ID))
                        ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                            Sqlstr = String.Format(String.Format("select COUNT(*) AS COUNT, PalletID, CartonID from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' GROUP BY PalletID, CartonID ", ID))
                        End If

                        dsCount = New DataSet
                        dsCount = da.ExecuteDataSet(Sqlstr, "IDInfor")
                    End If
                End If

                Return dsCount
            Catch ex As Exception
                ErrorLogging("MMC-CountOfBoxID", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Throw ex
            End Try
        End Using
    End Function

    Public Function get_uploadinfor(ByVal upload As DataSet, ByVal mydataset As DataSet, ByVal header As ProdPickingStructure, ByVal ERPLoginData As ERPLogin) As DataSet
        Dim i As Integer
        Dim QtyStr As Get_Qty
        Try

            Dim BOMData As DataTable
            Dim myDataColumn As DataColumn
            BOMData = New Data.DataTable("BOMData")


            myDataColumn = New Data.DataColumn("Material", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Revision", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("PerQty", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ManufacturerPN", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Manufacturer", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("ReqQty", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("AvlQty", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            myDataColumn = New Data.DataColumn("Shortage", System.Type.GetType("System.String"))
            BOMData.Columns.Add(myDataColumn)
            Dim myDataRow As Data.DataRow
            Dim qtydiff As Decimal
            Dim j As Integer
            Dim BuildQty As Integer = 0
            Dim n As Integer

            For i = 0 To upload.Tables(0).Rows.Count - 1
                myDataRow = BOMData.NewRow()
                myDataRow(0) = upload.Tables(0).Rows(i)(0)
                myDataRow(1) = ""
                myDataRow(2) = upload.Tables(0).Rows(i)(1)
                If Not upload.Tables(0).Rows(i).Item(2) Is DBNull.Value Then
                    myDataRow(3) = upload.Tables(0).Rows(i).Item(2)
                Else
                    myDataRow(3) = ""
                End If
                If Not upload.Tables(0).Rows(i).Item(3) Is DBNull.Value Then
                    myDataRow(4) = upload.Tables(0).Rows(i).Item(3)
                Else
                    myDataRow(4) = ""
                End If

                If upload.Tables(0).Rows(i)(0) <> "" Then
                    QtyStr = GetAvlQty(upload.Tables(0).Rows(i)(0), ERPLoginData)
                    myDataRow(6) = QtyStr.AvlQty
                End If

                BuildQty = 0
                For j = 0 To mydataset.Tables(0).Rows.Count - 1
                    BuildQty = BuildQty + mydataset.Tables(0).Rows(j)("IssueQty")
                Next

                If mydataset.Tables(0).Rows.Count > 1 Then
                    For n = 0 To mydataset.Tables(0).Rows.Count - 2
                        If mydataset.Tables(0).Rows(n)(1) = mydataset.Tables(0).Rows(n + 1)(1) Then
                            myDataRow(5) = myDataRow(2) * BuildQty
                        Else
                            myDataRow(5) = myDataRow(2) * header.BuildQty
                        End If
                    Next

                Else
                    myDataRow(5) = myDataRow(2) * BuildQty
                End If

                qtydiff = myDataRow(5) - myDataRow(6)
                myDataRow(7) = IIf(qtydiff > 0, qtydiff, 0)

                BOMData.Rows.Add(myDataRow)
            Next

            get_uploadinfor = New DataSet
            get_uploadinfor.Tables.Add(BOMData)
            Return get_uploadinfor

        Catch ex As Exception
            ErrorLogging("MMC-get_uploadinfor", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Throw ex
        End Try
    End Function

    Public Function CreateCLIDforPo(ByVal intUnit As Integer) As DataSet
        Dim myCLIDDataSet As DataSet
        Dim myCLIDTable As DataTable
        Dim i As Integer
        Dim myDataRow As DataRow
        Dim myDataColumn As DataColumn

        myCLIDTable = New DataTable("CLIDs")
        myDataColumn = New DataColumn("CLID", System.Type.GetType("System.String"))
        myCLIDTable.Columns.Add(myDataColumn)

        For i = 0 To intUnit - 1
            myDataRow = myCLIDTable.NewRow()
            myDataRow("CLID") = CreateNewCLID("PCBID")
            myCLIDTable.Rows.Add(myDataRow)
        Next
        myCLIDDataSet = New DataSet
        myCLIDDataSet.Tables.Add(myCLIDTable)
        myCLIDDataSet.AcceptChanges()
        Return myCLIDDataSet
    End Function

    Public Function CreateNewCLID(ByVal typeID As String) As String
        Dim NextCLID As String
        Dim ra, k As Integer
        Dim mycommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))

        mycommand = New SqlClient.SqlCommand("sp_Get_Next_Number", myConn)
        mycommand.CommandType = CommandType.StoredProcedure
        mycommand.CommandTimeout = TimeOut_M5
        mycommand.Parameters.AddWithValue("@NextNo", "")
        mycommand.Parameters(0).SqlDbType = SqlDbType.VarChar
        mycommand.Parameters(0).Size = 20
        mycommand.Parameters(0).Direction = ParameterDirection.Output
        'mycommand.Parameters.Add("@TypeID", typeID)
        mycommand.Parameters.AddWithValue("@TypeID", typeID)

        Try
            myConn.Open()

            'Try up to 5 times when failed getting next id
            NextCLID = ""
            k = 0
            While (k < 5 And NextCLID = "")
                Try
                    ra = mycommand.ExecuteNonQuery()
                    NextCLID = mycommand.Parameters(0).Value
                Catch ex As Exception
                    k = k + 1
                    ErrorLogging("MMC-CreateNewCLID", "Deadlocked? " & Str(k), "Failed getting next ID; typeID: " & typeID & ", " & ex.Message & ex.Source, "E")
                End Try
            End While

            myConn.Close()
            Return NextCLID

        Catch ex As Exception
            ErrorLogging("MMC-CreateNewCLID", "", "typeID: " & typeID & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function CountOfPoCLID(ByVal PO As String) As Double
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Object
        Dim strCMD As String
        strCMD = String.Format("exec sp_CountPOCLID '{0}'", PO)
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand(strCMD, myConn)
            cmdSQL.CommandTimeout = TimeOut_M5
            ra = cmdSQL.ExecuteScalar
            myConn.Close()
            Return Convert.ToDouble(ra)
        Catch ex As Exception
            ErrorLogging("MMC-CountOfPoCLID", "", "PO: " & PO & ", " & ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function SaveCreateCLID(ByVal ds As DataSet, ByVal ERPLoginData As ERPLogin, ByVal Productlabel As LabelData, ByVal Flag As Integer) As dj_response
        Using da As DataAccess = GetDataAccess()
            Dim strCMD As String
            Dim RTLot, CLID, boxID, OrgCode, matlNo, Description, Rohs, UOM, PurOrdNo, statusCode, LotNo, DateCode, Revision, CreatedBy, SubInv, Location, SONo, SOLine, MPN, MFR, LastTransaction As String
            Dim ExpDate As Date
            Dim QtyBaseUOM As Decimal
            Dim i As Integer
            Dim result1 As New Object
            Dim strCheck, CheckCLID, Msg As String
            Try
                Msg = "LabelID list: "
                If Flag = 1 Then
                    For i = 0 To ds.Tables("PoList").Rows.Count - 1
                        ds.Tables("PoList").Rows(i)("CLID") = CreateNewCLID("PCBID")
                        ds.Tables("PoList").AcceptChanges()
                        ds.AcceptChanges()

                        CLID = ds.Tables("PoList").Rows(i)("CLID")
                        If i < ds.Tables("PoList").Rows.Count - 1 Then
                            Msg = Msg & CLID & ", "
                        ElseIf i = ds.Tables("PoList").Rows.Count - 1 Then
                            Msg = Msg & CLID & "."
                        End If

                        boxID = Productlabel.LabelID.Trim
                        OrgCode = ERPLoginData.OrgCode
                        matlNo = ds.Tables("PoList").Rows(i)("Product")
                        Description = SQLString(Productlabel.ItemData.Tables(0).Rows(0)("item_desc"))
                        Rohs = Productlabel.RoHS
                        If Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE") = "" Then
                            UOM = "EA"
                        Else
                            UOM = Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE")
                        End If
                        PurOrdNo = ds.Tables("PoList").Rows(i)("PurOrdNo").ToString.Trim
                        QtyBaseUOM = ds.Tables("PoList").Rows(i)("QtyBaseUOM")
                        LotNo = Mid(SQLString(FixNull(Productlabel.LotNo)), 1, 50)

                        '' Adds Date Code field.
                        DateCode = Mid(SQLString(FixNull(Productlabel.DCode)), 1, 20)

                        statusCode = ds.Tables("PoList").Rows(i)("statusCode")
                        CreatedBy = ERPLoginData.User.ToUpper
                        Revision = Productlabel.MatlRev
                        SubInv = Productlabel.SubInv
                        Location = Productlabel.Locator
                        SONo = Productlabel.SONo
                        SOLine = Productlabel.SOLine
                        MPN = SQLString(Productlabel.MPN)
                        MFR = SQLString(Productlabel.MFR)
                        LastTransaction = "DJCompletion"        '"DJCompletion/MiscReceipt"
                        RTLot = Productlabel.RTLot
                        If Productlabel.ExpDate <> "" Then
                            ExpDate = Productlabel.ExpDate
                            strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS,DateCode,RTLot,ExpDate,Manufacturer,ManufacturerPN) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs, DateCode, RTLot, ExpDate, MFR, MPN)
                        Else
                            strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS,DateCode,RTLot, Manufacturer, ManufacturerPN) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}', '{20}','{21}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs, DateCode, RTLot, MFR, MPN)
                        End If
                        da.ExecuteNonQuery(strCMD)
                    Next

                ElseIf Flag = 2 Then
                    For i = 0 To ds.Tables("PoList").Rows.Count - 1
                        CLID = ds.Tables("PoList").Rows(i)("PurOrdNo")
                        If i < ds.Tables("PoList").Rows.Count - 1 Then
                            Msg = Msg & CLID & ", "
                        ElseIf i = ds.Tables("PoList").Rows.Count - 1 Then
                            Msg = Msg & CLID & "."
                        End If

                        OrgCode = ERPLoginData.OrgCode
                        matlNo = ds.Tables("PoList").Rows(i)("Product")
                        If Not ds.Tables("PoList").Rows(i)("CLID") Is DBNull.Value Then
                            boxID = ds.Tables("PoList").Rows(i)("CLID")
                        Else
                            boxID = ""
                        End If

                        If Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE") = "" Then
                            UOM = "EA"
                        Else
                            UOM = Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE")
                        End If

                        'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and (OrgCode is NULL or OrgCode = '{1}')", CLID, ERPLoginData.OrgCode))
                        'If result1.ToString <> "" Then
                        '    'If Not result1.ToString Is DBNull.Value Then
                        '    PurOrdNo = result1.ToString.Trim
                        'Else
                        '    PurOrdNo = ""
                        'End If
                        PurOrdNo = Productlabel.PurOrdNo

                        QtyBaseUOM = ds.Tables("PoList").Rows(i)("QtyBaseUOM")
                        LotNo = Mid(SQLString(FixNull(Productlabel.LotNo)), 1, 50)

                        '' Adds Date Code field.
                        DateCode = Mid(SQLString(FixNull(Productlabel.DCode)), 1, 20)

                        statusCode = ds.Tables("PoList").Rows(i)("statusCode")
                        CreatedBy = ERPLoginData.User.ToUpper
                        Revision = Productlabel.MatlRev
                        Description = SQLString(Productlabel.Description)
                        Rohs = Productlabel.RoHS
                        SubInv = Productlabel.SubInv
                        Location = Productlabel.Locator
                        SONo = Productlabel.SONo
                        SOLine = Productlabel.SOLine
                        MPN = SQLString(Productlabel.MPN)
                        MFR = SQLString(Productlabel.MFR)
                        LastTransaction = "DJCompletion"
                        strCheck = String.Format("select CLID from T_CLMaster where CLID = '{0}'", CLID)
                        CheckCLID = Convert.ToString(da.ExecuteScalar(strCheck))
                        RTLot = Productlabel.RTLot
                        If Productlabel.ExpDate <> "" Then
                            If CheckCLID = "" Then
                                strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS,DateCode,RTLot,ExpDate,Manufacturer,ManufacturerPN) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs, DateCode, RTLot, ExpDate, MFR, MPN)
                            Else
                                strCMD = String.Format("UPDATE T_CLMaster set OrgCode='{0}',MaterialNo='{1}',MaterialRevision='{2}',QtyBaseUOM='{3}',BaseUOM='{4}',RecDate=getDate(),ChangedOn=getDate(), ChangedBy='{5}', StatusCode='1', LotNo='{6}', SLOC='{7}', StorageBin='{8}', SONo='{9}', SOLine='{10}', PurOrdNo='{11}',LastTransaction='{12}',DateCode='{13}',RTLot='{14}',ExpDate='{15}',Manufacturer='{16}',ManufacturerPN='{17}',BoxID='{18}' where CLID='{19}'", OrgCode, matlNo, Revision, QtyBaseUOM, UOM, ERPLoginData.User, LotNo, SubInv, Location, SONo, SOLine, PurOrdNo, LastTransaction, DateCode, RTLot, ExpDate, MFR, MPN, boxID, CLID)
                            End If
                        Else
                            If CheckCLID = "" Then
                                strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS,DateCode,RTLot,Manufacturer,ManufacturerPN) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs, DateCode, RTLot, MFR, MPN)
                            Else
                                strCMD = String.Format("UPDATE T_CLMaster set OrgCode='{0}',MaterialNo='{1}',MaterialRevision='{2}',QtyBaseUOM='{3}',BaseUOM='{4}',RecDate=getDate(),ChangedOn=getDate(), ChangedBy='{5}', StatusCode='1', LotNo='{6}', SLOC='{7}', StorageBin='{8}', SONo='{9}', SOLine='{10}', PurOrdNo='{11}',LastTransaction='{12}',DateCode='{13}',RTLot='{14}',Manufacturer='{15}',ManufacturerPN='{16}',BoxID='{17}' where CLID='{18}'", OrgCode, matlNo, Revision, QtyBaseUOM, UOM, ERPLoginData.User, LotNo, SubInv, Location, SONo, SOLine, PurOrdNo, LastTransaction, DateCode, RTLot, MFR, MPN, boxID, CLID)
                            End If
                        End If
                        da.ExecuteNonQuery(strCMD)
                        'ErrorLogging("MaterialReplenishment-SaveCreateCLID-Finish", ERPLoginData.User, "DJ: " & PurOrdNo & " ,BoxID: " & CLID, "I")
                    Next
                End If
                ErrorLogging("MaterialReplenishment-SaveCreateCLID-Finish", ERPLoginData.User, Msg, "I")
                SaveCreateCLID.dsCLID = ds
                SaveCreateCLID.flag = "Y"
                SaveCreateCLID.errormsg = ""
                Return SaveCreateCLID
            Catch ex As Exception
                SaveCreateCLID.dsCLID = ds
                SaveCreateCLID.flag = "N"
                SaveCreateCLID.errormsg = ex.Message.ToString
                ErrorLogging("MaterialReplenishment-SaveCreateCLID", "", "DJ: " & PurOrdNo & ", " & Msg & ", Error: " & ex.Message & ex.Source, "E")
                Return SaveCreateCLID
            End Try
        End Using
    End Function

    Public Function SaveCreatCLIDforPo(ByVal ds As DataSet, ByVal SAPLoginData As ERPLogin, ByVal Productlabel As LabelData, ByVal Flag As Integer)
        Dim CLMasterSQLCommand As SqlClient.SqlCommand
        Dim CheckCommand As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim ra As Integer
        Dim strCMD As String
        Dim CLID, boxID, OrgCode, matlNo, Description, Rohs, UOM, PurOrdNo, statusCode, LotNo, Revision, CreatedBy, SubInv, Location, SONo, SOLine, LastTransaction As String
        Dim QtyBaseUOM As Decimal
        Dim i As Integer
        Dim result1 As New Object
        Dim DJNO As String
        Dim strCheck, CheckCLID As String

        If Flag = 1 Then
            For i = 0 To ds.Tables("PoList").Rows.Count - 1
                CLID = ds.Tables("PoList").Rows(i)("CLID")
                'boxID = ProductLabel.LabelID.Trim
                OrgCode = SAPLoginData.OrgCode
                matlNo = ds.Tables("PoList").Rows(i)("Product")
                Description = SQLString(Productlabel.ItemData.Tables(0).Rows(0)("item_desc"))
                Rohs = Productlabel.RoHS
                If Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE") = "" Then
                    UOM = "EA"
                Else
                    UOM = Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE")
                End If
                PurOrdNo = ds.Tables("PoList").Rows(i)("PurOrdNo")
                QtyBaseUOM = ds.Tables("PoList").Rows(i)("QtyBaseUOM")
                LotNo = ds.Tables("PoList").Rows(i)("LotNo")
                statusCode = ds.Tables("PoList").Rows(i)("statusCode")
                CreatedBy = SAPLoginData.User.ToUpper
                Revision = Productlabel.MatlRev
                SubInv = Productlabel.SubInv
                Location = Productlabel.Locator
                SONo = Productlabel.SONo
                SOLine = Productlabel.SOLine
                '2010-10-7
                LastTransaction = "DJCompletion/MiscReceipt"
                'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,MaterialRevision,SLOC,StorageBin) values ('{0}','{1}','{2}',0,'{3}',{4},'{5}','{6}','EA',getDate(),getDate(),'{7}','{8}','{9}','{10}')", CLID, OrgCode, matlNo, PurOrdNo, QtyBaseUOM, LotNo, statusCode, CreatedBy, Revision, SubInv, Location)

                'strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs)
                strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,SONo,SOLine,LastTransaction,MaterialDesc,RoHS) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, SONo, SOLine, LastTransaction, Description, Rohs)


                Try
                    myConn.Open()
                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                    ra = CLMasterSQLCommand.ExecuteNonQuery()
                    myConn.Close()
                Catch ex As Exception
                    ErrorLogging("MMC-SaveCreatCLIDforPo", SAPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try
            Next
        Else
            For i = 0 To ds.Tables("PoList").Rows.Count - 1
                CLID = ds.Tables("PoList").Rows(i)("PurOrdNo")
                OrgCode = SAPLoginData.OrgCode
                matlNo = ds.Tables("PoList").Rows(i)("Product")
                If Not ds.Tables("PoList").Rows(i)("CLID") Is DBNull.Value Then
                    boxID = ds.Tables("PoList").Rows(i)("CLID")
                Else
                    boxID = ""
                End If

                If Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE") = "" Then
                    UOM = "EA"
                Else
                    UOM = Productlabel.ItemData.Tables(0).Rows(0)("UOM_CODE")
                End If
                Using da As DataAccess = GetDataAccess()
                    'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and (OrgCode is NULL or OrgCode = '{1}')", CLID, SAPLoginData.OrgCode))
                    result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}'", CLID))

                    If result1.ToString <> "" Then
                        'If Not result1.ToString Is DBNull.Value Then
                        PurOrdNo = result1.ToString
                    Else
                        PurOrdNo = ""
                    End If
                End Using
                QtyBaseUOM = ds.Tables("PoList").Rows(i)("QtyBaseUOM")
                LotNo = ds.Tables("PoList").Rows(i)("LotNo")
                statusCode = ds.Tables("PoList").Rows(i)("statusCode")
                CreatedBy = SAPLoginData.User.ToUpper
                Revision = Productlabel.MatlRev
                Description = SQLString(Productlabel.Description)
                Rohs = Productlabel.RoHS
                SubInv = Productlabel.SubInv
                Location = Productlabel.Locator
                SONo = Productlabel.SONo
                SOLine = Productlabel.SOLine
                '2010-10-7
                LastTransaction = "DJCompletion"
                Try
                    myConn.Open()
                    strCheck = String.Format("select CLID from T_CLMaster where CLID = '{0}'", CLID)
                    CheckCommand = New SqlClient.SqlCommand(strCheck, myConn)
                    CheckCLID = Convert.ToString(CheckCommand.ExecuteScalar)
                    If CheckCLID = "" Then
                        strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, QtyBaseUOM, LotNo, statusCode, UOM, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, Rohs)
                    Else
                        strCMD = String.Format("UPDATE T_CLMaster set OrgCode='{0}',MaterialNo='{1}',MaterialRevision='{2}',QtyBaseUOM='{3}',BaseUOM='{4}',RecDate=getDate(),ChangedOn=getDate(), ChangedBy='{5}', StatusCode='1', LotNo='{6}', SLOC='{7}', StorageBin='{8}', SONo='{9}', SOLine='{10}', PurOrdNo='{11}',LastTransaction='{12}' where CLID='{13}'", OrgCode, matlNo, Revision, QtyBaseUOM, UOM, SAPLoginData.User, LotNo, SubInv, Location, SONo, SOLine, PurOrdNo, LastTransaction, CLID)
                    End If

                    CLMasterSQLCommand = New SqlClient.SqlCommand(strCMD, myConn)
                    CLMasterSQLCommand.CommandTimeout = TimeOut_M5
                    ra = CLMasterSQLCommand.ExecuteNonQuery()
                    myConn.Close()
                Catch ex As Exception
                    ErrorLogging("MMC-SaveCreatCLIDforPo", SAPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
                Finally
                    If myConn.State <> ConnectionState.Closed Then myConn.Close()
                End Try
            Next
        End If
    End Function

    Public Function RePrintCreateCLIDforPOCLIDs(ByVal CLIDs As DataSet, ByVal Printer As String) As Boolean
        'RePrint for normal CLIDs
        RePrintCreateCLIDforPOCLIDs = True
        Try
            'Create instance of Codesoft and open label file
            'app1 = New LabelManager2.Application
            'app1.Documents.Open(LabelFile)

            Dim i As Integer
            For i = 0 To CLIDs.Tables(0).Rows.Count - 1
                'If PrintCreateCLIDforPOCLID(CLIDs.Tables(0).Rows(i)(0), Printer) = False Then
                If PrintCLID(CLIDs.Tables(0).Rows(i)("CLID"), Printer) = False Then
                    RePrintCreateCLIDforPOCLIDs = False
                End If
            Next

            'Quit Code Soft...
            'app1.Documents.CloseAll(False)
            'app1.Quit()

        Catch ex As Exception
            ErrorLogging("MMC-RePrintCreateCLIDforPOCLIDs", "", ex.Message & ex.Source, "E")
        End Try

    End Function

    Public Function Post_DJ_Completion(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String, ByVal ds As DataSet, ByVal Productlabel As LabelData, ByVal Flag As Integer) As dj_response
        Using da As DataAccess = GetDataAccess()
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator, "I")
            Dim aa As OracleString
            Dim bb As Int32
            Dim uom As String
            Dim rev As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()

            Try
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                'If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4) Is DBNull.Value Then
                '    uom = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4)
                'Else
                '    uom = ""
                'End If
                'If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1) Is DBNull.Value And Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(10) Is DBNull.Value Then
                '    rev = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1)
                'Else
                '    rev = ""
                'End If
                uom = Productlabel.UOM
                rev = Productlabel.MatlRev

                DJ = DJ.Trim

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_complete"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'ERPLoginData.OrgCode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_complete_qty", OracleType.Number).Value = com_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = rev
                Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Value = Inventory
                Oda.Parameters.Add("x_locator", OracleType.VarChar, 150).Value = Locator

                Oda.Parameters("x_subinventory").Direction = ParameterDirection.InputOutput
                Oda.Parameters("x_locator").Direction = ParameterDirection.InputOutput

                Oda.Parameters.Add("o_lot_num", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_expiration_date", OracleType.DateTime).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                bb = Oda.ExecuteOracleNonQuery(aa)
                Post_DJ_Completion.dsCLID = ds
                Post_DJ_Completion.flag = Oda.Parameters("o_success_flag").Value.ToString()
                Post_DJ_Completion.errormsg = Oda.Parameters("o_error_msgg").Value.ToString()
                ErrorLogging("MaterialReplenishment-Post_DJ_Completion", ERPLoginData.User, "DJ: " & DJ & " , Qty: " & com_qty & " , PostType: " & Flag & " , ReturnFlag: " & Post_DJ_Completion.flag & " , ReturnMsg: " & Post_DJ_Completion.errormsg, "I")
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()
                    Productlabel.RTLot = FixNull(Oda.Parameters("o_lot_num").Value)
                    Productlabel.ExpDate = FixNull(Oda.Parameters("o_expiration_date").Value)
                    'ErrorLogging("MaterialReplenishment-SaveCreateCLID-Start", ERPLoginData.User, "DJ: " & DJ, "I")
                    Post_DJ_Completion = SaveCreateCLID(ds, ERPLoginData, Productlabel, Flag)
                Else
                    Oda.Transaction.Rollback()
                End If
                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return Post_DJ_Completion
            Catch ex As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("MaterialReplenishment-Post_DJ_Completion", ERPLoginData.User, "DJ: " & DJ & ", " & ex.Message & ex.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    Oda.Connection.Dispose()
                End If
                Post_DJ_Completion.flag = "N"
                Post_DJ_Completion.errormsg = ex.Message.ToString
                Post_DJ_Completion.dsCLID = ds
                Return Post_DJ_Completion
            End Try
        End Using
    End Function

    Public Function DJ_Completion(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String) As dj_response
        Using da As DataAccess = GetDataAccess()
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator, "I")
            Dim aa As OracleString
            Dim bb As Int32
            Dim uom As String
            Dim rev As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()
            Try
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4) Is DBNull.Value Then
                    uom = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4)
                Else
                    uom = ""
                End If
                If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1) Is DBNull.Value And Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(10) Is DBNull.Value Then
                    rev = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1)
                Else
                    rev = ""
                End If

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_complete"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'ERPLoginData.OrgCode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_complete_qty", OracleType.Number, 50).Value = com_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = rev
                Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Value = Inventory
                Oda.Parameters.Add("x_locator", OracleType.VarChar, 100).Value = Locator
                'Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Direction = ParameterDirection.InputOutput
                'Oda.Parameters.Add("x_locator", OracleType.VarChar, 100).Direction = ParameterDirection.InputOutput
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output

                bb = Oda.ExecuteOracleNonQuery(aa)
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()
                Else
                    Oda.Transaction.Rollback()
                End If
                DJ_Completion.flag = Oda.Parameters("o_success_flag").Value.ToString()
                DJ_Completion.errormsg = Oda.Parameters("o_error_msgg").Value.ToString()
                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                'Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return DJ_Completion
            Catch oe As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    ' Oda.Connection.Dispose()
                End If
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try

        End Using
    End Function

    Public Function DJ_Completion_BoxID(ByVal ERPLoginData As ERPLogin, ByVal BoxID As String, ByVal com_qty As Integer, ByVal Inventory As String, ByVal Locator As String) As dj_response
        Using da As DataAccess = GetDataAccess()
            Dim aa As OracleString
            Dim bb As Int32
            Dim uom As String
            Dim rev As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            Dim DJ As String
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()

            Dim result1 As New Object

            If Mid(BoxID, 1, 1) <> "P" AndAlso Mid(BoxID, 1, 1) <> "1" Then
                'ID.Contains("B") Then
                'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}' and (OrgCode is NULL or OrgCode = '{1}')", BoxID, ERPLoginData.OrgCode))
                result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where CartonID = '{0}'", BoxID))

            ElseIf Mid(BoxID, 1, 1) = "P" OrElse Mid(BoxID, 1, 1) = "1" Then
                'ID.Contains("P") Then
                'result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}' and (OrgCode is NULL or OrgCode = '{1}')", BoxID, ERPLoginData.OrgCode))
                result1 = da.ExecuteScalar(String.Format("select distinct ProdOrder from T_Shippment with (nolock) where PalletID = '{0}'", BoxID))

            End If

            If result1.ToString <> "" Then
                'If Not result1.ToString Is DBNull.Value Then
                DJ = result1.ToString
            Else
                DJ = ""
            End If

            If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4) Is DBNull.Value Then
                uom = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(4)
            Else
                uom = ""
            End If

            If Not GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1) Is DBNull.Value Then
                rev = GetOrderInfoFromERP(ERPLoginData, DJ).Tables(0).Rows(0).Item(1)
            Else
                rev = ""
            End If
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator, "I")
            Try
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_complete"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'ERPLoginData.OrgCode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_complete_qty", OracleType.Number, 50).Value = com_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = rev
                Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Value = Inventory
                Oda.Parameters.Add("x_locator", OracleType.VarChar, 100).Value = Locator
                'Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Direction = ParameterDirection.InputOutput
                'Oda.Parameters.Add("x_locator", OracleType.VarChar, 100).Direction = ParameterDirection.InputOutput
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 2000).Direction = ParameterDirection.Output

                bb = Oda.ExecuteOracleNonQuery(aa)
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()
                Else
                    Oda.Transaction.Rollback()
                End If
                DJ_Completion_BoxID.flag = Oda.Parameters("o_success_flag").Value.ToString()
                DJ_Completion_BoxID.errormsg = Oda.Parameters("o_error_msgg").Value.ToString()
                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                'Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return DJ_Completion_BoxID
            Catch oe As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("MaterialReplenishment-DJ_Completion_BoxID", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    'Oda.Connection.Dispose()
                End If
            Finally
                If Oda.Connection.State <> ConnectionState.Closed Then Oda.Connection.Close()
            End Try

        End Using
    End Function

    Public Function DJCompletion_For_PMJob(ByVal ERPLoginData As ERPLogin, ByVal DJ As String, ByVal dj_rev As String, ByVal Assembly As String, ByVal com_qty As String, ByVal uom As String, ByVal SubInventory As String, ByVal Locator As String, ByVal ROHS As String, ByVal MSB As String, ByVal item_desc As String, ByVal LabelPrinter As String) As String
        Using da As DataAccess = GetDataAccess()
            'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & "com_qty: " & com_qty & "Sub Inventory: " & Inventory & "Locator: " & Locator, "I")
            Dim aa As OracleString
            Dim bb As Int32
            Dim strCMD, flag, Msg, RTLot, ExpDate, CLID, boxID, OrgCode, matlNo, Description, PurOrdNo, statusCode, LotNo, DateCode, Revision, CreatedBy, SubInv, Location, SONo, SOLine, MPN, MFR, LastTransaction As String
            Dim resp As Integer = 20560    '54050
            Dim appl As Integer = 706
            'rev = ""
            Dim Oda As OracleCommand = da.Ora_Command_Trans()

            Try
                Dim OrgID As String = GetOrgID(ERPLoginData.OrgCode)

                DJ = DJ.Trim

                If Oda.Connection.State = ConnectionState.Closed Then
                    Oda.Connection.Open()
                End If

                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.initialize"
                Oda.Parameters.Add("p_user_id", OracleType.Int32).Value = CInt(ERPLoginData.UserID) '15904
                Oda.Parameters.Add("p_resp_id", OracleType.Int32).Value = ERPLoginData.RespID_WIP  '54050
                Oda.Parameters.Add("p_appl_id", OracleType.Int32).Value = ERPLoginData.AppID_WIP
                Oda.ExecuteOracleNonQuery(aa)
                Oda.Parameters.Clear()


                Oda.CommandType = CommandType.StoredProcedure
                Oda.CommandText = "apps.XXETR_wip_pkg.process_wip_complete"
                Oda.Parameters.Add("p_dj_name", OracleType.VarChar, 50).Value = DJ
                Oda.Parameters.Add("p_organization_code", OracleType.VarChar, 100).Value = OrgID  'ERPLoginData.OrgCode   'fixed Org Code 404 for Test only
                Oda.Parameters.Add("p_complete_qty", OracleType.Number).Value = com_qty
                Oda.Parameters.Add("p_uom", OracleType.VarChar, 50).Value = uom
                Oda.Parameters.Add("p_rev", OracleType.VarChar, 10).Value = ""
                Oda.Parameters.Add("x_subinventory", OracleType.VarChar, 150).Value = SubInventory
                Oda.Parameters.Add("x_locator", OracleType.VarChar, 150).Value = Locator

                Oda.Parameters("x_subinventory").Direction = ParameterDirection.InputOutput
                Oda.Parameters("x_locator").Direction = ParameterDirection.InputOutput

                Oda.Parameters.Add("o_lot_num", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_expiration_date", OracleType.DateTime).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_success_flag", OracleType.VarChar, 50).Direction = ParameterDirection.Output
                Oda.Parameters.Add("o_error_msgg", OracleType.VarChar, 5000).Direction = ParameterDirection.Output

                bb = Oda.ExecuteOracleNonQuery(aa)
                flag = Oda.Parameters("o_success_flag").Value.ToString()
                DJCompletion_For_PMJob = Oda.Parameters("o_error_msgg").Value.ToString()

                ErrorLogging("EPM-Post_DJ_Completion", ERPLoginData.User, "DJ: " & DJ & " , Qty: " & com_qty & " , PostType: " & " , ReturnFlag: " & flag & " , ReturnMsg: " & DJCompletion_For_PMJob, "I")
                If Oda.Parameters("o_success_flag").Value.ToString() = "Y" Then
                    Oda.Transaction.Commit()
                    RTLot = FixNull(Oda.Parameters("o_lot_num").Value)
                    ExpDate = FixNull(Oda.Parameters("o_expiration_date").Value)
                    'ErrorLogging("MaterialReplenishment-SaveCreateCLID-Start", ERPLoginData.User, "DJ: " & DJ, "I")
                    CLID = CreateNewCLID("PCBID")
                    Msg = "LabelID list: " & CLID & "."

                    boxID = ""
                    OrgCode = ERPLoginData.OrgCode
                    matlNo = Assembly
                    Description = item_desc
                    ROHS = ROHS
                    PurOrdNo = DJ

                    Dim Flag_AutoRev As String
                    Flag_AutoRev = Flag_AutoGetBOMRev(ERPLoginData)
                    If Flag_AutoRev.ToUpper = "YES" Then
                        LotNo = FixNull(dj_rev)
                        If LotNo = "" Then
                            Dim itemds As New DataSet
                            itemds = Get_Item_Master(ERPLoginData, FixNull(Assembly))
                            If Not itemds Is Nothing AndAlso itemds.Tables.Count > 0 AndAlso itemds.Tables(0).Rows.Count > 0 Then
                                LotNo = FixNull(itemds.Tables(0).Rows(0)("item_rev"))
                            End If
                        End If
                    End If

                    LotNo = Mid(SQLString(FixNull(LotNo)), 1, 50)

                    '' Adds Date Code field.
                    DateCode = ""

                    statusCode = "1"
                    CreatedBy = ERPLoginData.User.ToUpper
                    Revision = ""
                    SubInv = SubInventory
                    Location = Locator
                    SONo = ""
                    SOLine = ""
                    MPN = MSB
                    MFR = MSB
                    LastTransaction = "DJCompletion"        '"DJCompletion/MiscReceipt"
                    RTLot = ""

                    strCMD = String.Format("INSERT INTO T_CLMaster (CLID,OrgCode,MaterialNo,MaterialRevision,Qty,PurOrdNo,QtyBaseUOM,LotNo,statusCode,BaseUOM,RecDate,CreatedOn,CreatedBy,SLOC,StorageBin,BoxID,SONo,SOLine,LastTransaction,MaterialDesc,RoHS,DateCode,RTLot, Manufacturer, ManufacturerPN) values ('{0}','{1}','{2}','{3}',0,'{4}',{5},'{6}','{7}','{8}',getDate(),getDate(),'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}', '{20}','{21}')", CLID, OrgCode, matlNo, Revision, PurOrdNo, com_qty, LotNo, statusCode, uom, CreatedBy, SubInv, Location, boxID, SONo, SOLine, LastTransaction, Description, ROHS, DateCode, RTLot, MFR, MPN)
                    da.ExecuteNonQuery(strCMD)

                    PrintCLID(CLID, LabelPrinter)
                Else
                    Oda.Transaction.Rollback()
                End If
                'DJ_Completion.subInv = Oda.Parameters("x_subinventory").Value.ToString()
                'DJ_Completion.locator = Oda.Parameters("x_locator").Value.ToString()

                Oda.Connection.Close()
                Oda.Connection.Dispose()

                'Return DirectCast(DJFlag, String)
                Return DJCompletion_For_PMJob
            Catch ex As Exception
                'ErrorLogging("MaterialReplenishment-DJ_Completion", "", "DJ: " & DJ & ", " & oe.Message & oe.Source, "E")
                ErrorLogging("EPM-DJCompletion_For_PMJob", ERPLoginData.User, "DJ: " & DJ & ", " & ex.Message & ex.Source, "E")
                If Oda.Connection.State = ConnectionState.Open Then
                    Oda.Transaction.Rollback()
                    Oda.Connection.Close()
                    Oda.Connection.Dispose()
                End If
                flag = "N"
                DJCompletion_For_PMJob = ex.Message.ToString
                Return DJCompletion_For_PMJob
            End Try
        End Using
    End Function

    Public Function Get_Item_Master(ByVal LoginData As ERPLogin, ByVal PartNo As String) As DataSet
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

    ''' <summary>
    ''' Print PO CLID.
    ''' </summary>
    ''' <param name="CLID"></param>
    ''' <param name="Printer"></param>
    ''' <returns></returns>
    ''' <remarks> 
    ''' Add LotNo field. 06-29-2011
    ''' Add DateCode field.
    ''' </remarks>
    ''' <modify> By Jackson Huang</modify>
    ''' <Date> 12-28-2011 </Date>
    ''' 
    Public Function PrintCreateCLIDforPOCLID(ByVal CLID As String, ByVal Printer As String) As Boolean
        'Print for normal CLID
        Dim NewCLIDCommand As SqlClient.SqlCommand
        Dim ra As Integer
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Dim objReader As SqlClient.SqlDataReader
        Dim lblData As POLabel
        Dim LabelPrinter, lblPrint As String
        Dim TmpDate As Date

        'LabelPrinter = app1.ActiveDocument.Printer.Name
        LabelPrinter = Printer
        PrintCreateCLIDforPOCLID = True

        NewCLIDCommand = New SqlClient.SqlCommand("Select CLID,OrgCode,MaterialNo,MaterialRevision,purOrdNo,QtyBaseUOM,BaseUOM,RecDate,SONo,SOLine,LotNo,DateCode,MPN from T_CLMaster with (nolock) where CLID = @CLID ", myConn)
        NewCLIDCommand.Parameters.Add("@CLID", SqlDbType.VarChar, 20, "CLID")
        NewCLIDCommand.Parameters("@CLID").Value = CLID

        Try
            myConn.Open()
            NewCLIDCommand.CommandTimeout = TimeOut_M5
            objReader = NewCLIDCommand.ExecuteReader()

            'app1 = New LabelManager2.Application
            'app1.Documents.Open(LabelforPOFile)
            While objReader.Read()
                If Not objReader.GetValue(0) Is DBNull.Value Then lblData.CLID = objReader.GetValue(0)
                If Not objReader.GetValue(1) Is DBNull.Value Then lblData.OrgCode = objReader.GetValue(1)
                If Not objReader.GetValue(2) Is DBNull.Value Then lblData.PCBA = objReader.GetValue(2)
                If Not objReader.GetValue(3) Is DBNull.Value Then lblData.Rev = objReader.GetValue(3)
                If Not objReader.GetValue(4) Is DBNull.Value Then lblData.POID = objReader.GetValue(4)
                If Not objReader.GetValue(5) Is DBNull.Value Then lblData.Qty = objReader.GetValue(5)
                If Not objReader.GetValue(6) Is DBNull.Value Then lblData.UOM = objReader.GetValue(6)
                If Not objReader.GetValue(7) Is DBNull.Value Then
                    TmpDate = objReader.GetValue(7)
                    lblData.RecDate = TmpDate.ToString("MM/dd/yyyy")
                End If
                If Not objReader.GetValue(8) Is DBNull.Value Then lblData.SONo = objReader.GetValue(8)
                If Not objReader.GetValue(9) Is DBNull.Value Then lblData.SOLine = objReader.GetValue(9)
                If Not objReader.GetValue(10) Is DBNull.Value Then lblData.LotNo = objReader.GetValue(10)
                '' Add DateCode field.
                If Not objReader.GetValue(11) Is DBNull.Value Then lblData.DateCode = objReader.GetValue(11)
                If Not objReader.GetValue(12) Is DBNull.Value Then lblData.MPN = objReader.GetValue(12)
                'lblPrint = PrintCreateCLIDforPOLabel(LabelPrinter, lblData)
                lblPrint = PrintPOLabel(LabelPrinter, lblData)

                'Fill variables
                'app1.ActiveDocument.Variables.Item("CLID").Value = lblData.CLID
                'app1.ActiveDocument.Variables.Item("PCBA").Value = lblData.PCBA
                'app1.ActiveDocument.Variables.Item("POID").Value = lblData.POID
                'app1.ActiveDocument.Variables.Item("Qty").Value = lblData.Qty
                'app1.ActiveDocument.Variables.Item("UOM").Value = lblData.UOM
                'app1.ActiveDocument.Variables.Item("RecDate").Value = lblData.RecDate
                ''Change printer if needed
                'If app1.ActiveDocument.Printer.Name <> Printer Then
                '    app1.ActiveDocument.Printer.SwitchTo(Printer, , False)
                'End If
                'app1.ActiveDocument.PrintDocument()
                'Quit Code Soft...
            End While
            'app1.Documents.CloseAll(False)
            'app1.Quit()
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("MMC-PrintCreateCLIDforPOCLID", "", "CLID: " & CLID & ", " & ex.Message & ex.Source, "E")
            PrintCreateCLIDforPOCLID = False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try

    End Function

    Public Function getStorageBin(ByVal CLID As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim Sqlstr As String
            Sqlstr = String.Format("SELECT StorageBin from T_CLMaster with (nolock) where CLID = '{0}'", CLID)
            Return Convert.ToString(da.ExecuteScalar(Sqlstr))
        End Using
    End Function

    Public Function getClidListForPick(ByVal matl As String, ByVal matlRev As String, ByVal ReqQty As Double, ByVal SubIn As String) As DataSet
        Using da As DataAccess = GetDataAccess()
            Return da.ExecuteDataSet(String.Format("exec sp_Get_ClidList_ForPick '{0}','{1}',{2},'{3}'", matl, matlRev, ReqQty, SubIn))
        End Using
    End Function

    '2010-8-16 For DJ Com
    Public Function GetSOLine(ByVal LoginData As ERPLogin, ByVal SONo As String, ByVal SOLine As String) As dj_response
        Dim lineNum, shipNum As String
        Dim orderLine() As String
        If SOLine = "" Then
            lineNum = ""
            shipNum = ""
        Else
            orderLine = Split(SOLine, ".")
            lineNum = orderLine(0)
            shipNum = orderLine(1)
        End If
        Using da As DataAccess = GetDataAccess()
            Dim oda As OracleDataAdapter = da.Oda_Sele()

            Try
                Dim ds As New DataSet()
                ds.Tables.Add("SOData")

                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_oe_pkg.get_so_item"
                oda.SelectCommand.Parameters.Add("p_order_num", OracleType.VarChar).Value = SONo
                oda.SelectCommand.Parameters.Add("p_line_num", OracleType.VarChar).Value = lineNum
                oda.SelectCommand.Parameters.Add("p_shipment_num", OracleType.VarChar).Value = shipNum
                oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output

                oda.SelectCommand.Connection.Open()
                oda.Fill(ds, "SOData")
                oda.SelectCommand.Connection.Close()

                GetSOLine.flag = ""
                GetSOLine.errormsg = ""
                GetSOLine.dsCLID = ds
                Return GetSOLine

            Catch oe As Exception
                ErrorLogging("DJ Completion-GetSOLine", LoginData.User, "SONo/SOLine: " & SONo & "/" & SOLine & ", " & oe.Message & oe.Source, "E")
                GetSOLine.errormsg = oe.Message.ToString.Trim
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using

    End Function

    'Getrelease_lines
    Public Function Getrelease_lines(ByVal ERPLoginData As ERPLogin, ByVal OrderNo As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim ds As New DataSet
            ds.Tables.Add("Pickorder")

            Dim AssCLIDList As DataTable
            Dim myDataColumn As DataColumn
            Dim ds2 As DataSet = New DataSet

            If ds2.Tables.Count = 0 Then
                AssCLIDList = New DataTable("Pickorder2")
                myDataColumn = New DataColumn("DestSubInv", System.Type.GetType("System.String"))
                AssCLIDList.Columns.Add(myDataColumn)
                myDataColumn = New DataColumn("DestLoctor", System.Type.GetType("System.String"))
                AssCLIDList.Columns.Add(myDataColumn)
                ds2.Tables.Add(AssCLIDList)
            End If

            Dim myDataRow2 As Data.DataRow
            Dim i As Integer

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim sda As SqlClient.SqlDataAdapter

            Try

                'oda.SelectCommand.CommandType = CommandType.StoredProcedure
                'oda.SelectCommand.CommandText = "APPS.XXETR_wip_pkg.get_release_lines"

                'oda.SelectCommand.Parameters.Add("o_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                'oda.SelectCommand.Parameters.Add("p_discreate_job", OracleType.VarChar).Value = OrderNo
                'oda.SelectCommand.Parameters.Add(New OracleParameter("p_org_code", OracleType.VarChar, 240)).Value = ERPLoginData.OrgCode

                'oda.SelectCommand.Connection.Open()
                'oda.Fill(ds, "Pickorder")
                'oda.SelectCommand.Connection.Close()

                sda = da.Sda_Sele()
                sda.SelectCommand.CommandType = CommandType.StoredProcedure
                sda.SelectCommand.CommandText = "dbo.ora_get_release_lines"
                sda.SelectCommand.CommandTimeout = TimeOut_M5

                sda.SelectCommand.Parameters.Add("@p_discreate_job", SqlDbType.VarChar, 150).Value = OrderNo
                sda.SelectCommand.Parameters.Add("@p_org_id", SqlDbType.Int, 150).Value = ERPLoginData.OrgID
                sda.SelectCommand.Connection.Open()
                sda.Fill(ds, "Pickorder")
                sda.SelectCommand.Connection.Close()


                For i = 0 To ds.Tables(0).Rows.Count - 1
                    'If ds.Tables(0).Rows(i)(9) = "FR0 BF1" Or ds.Tables(0).Rows(i)(9) = "FR0 BF2" Then
                    If InStr(ds.Tables(0).Rows(i)("supply_subinventory").ToString, "FS") > 0 OrElse InStr(UCase(ds.Tables(0).Rows(i)("sypply_type")), "PUSH") > 0 Then
                        Continue For
                    End If

                    If (FixNull(ds.Tables(0).Rows(i).Item("ejit_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("ejit_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("kanban_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("kanban_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("consumables_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("consumables_flag")) = "") AndAlso (FixNull(ds.Tables(0).Rows(i).Item("bulky_flag")) = "N" OrElse FixNull(ds.Tables(0).Rows(i).Item("bulky_flag")) = "") Then
                        myDataRow2 = ds2.Tables(0).NewRow()
                        myDataRow2(0) = ds.Tables(0).Rows(i)(9).ToString
                        myDataRow2(1) = ds.Tables(0).Rows(i)(10).ToString
                        ds2.Tables(0).Rows.Add(myDataRow2)
                        Exit For
                    End If
                Next

                Return ds2

            Catch ex As Exception
                ErrorLogging("MMC-Getrelease_lines", ERPLoginData.User.ToUpper, ex.Message & ex.Source, "E")
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
            Return ds2
        End Using

    End Function

    Public Function writeClidRefPo(ByVal listItem As DataSet, ByVal DjNo As String, ByVal userName As String) As Boolean
        Dim ClidQty As String = ""
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            For i As Integer = 0 To listItem.Tables("MatList").Rows.Count - 1
                Dim clid As String = listItem.Tables("MatList").Rows(i)("CLID").ToString()
                Dim qty As Decimal = Convert.ToDecimal(listItem.Tables("MatList").Rows(i)("QTY"))
                If ClidQty = "" Then
                    ClidQty = clid & "*" & qty
                Else
                    ClidQty = ClidQty & "," & clid & "*" & qty
                End If
            Next
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_ClidRefPo", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@v_DjNo", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@v_DjNo").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@v_DjNo").Value = DjNo

            cmdSQL.Parameters.Add("@v_ClidQty", SqlDbType.VarChar, 2000)
            cmdSQL.Parameters("@v_ClidQty").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@v_ClidQty").Value = ClidQty

            cmdSQL.Parameters.Add("@v_ChangeBy", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@v_ChangeBy").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@v_ChangeBy").Value = userName
            If cmdSQL.ExecuteScalar().ToString = "1" Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            ErrorLogging("MaterialReplenishment-writeClidRefPo", "", "DJNO: " & DjNo & ", " & ex.Message & ex.Source, "E")
            Return False
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
    End Function

    Public Function AllowEditDJ(ByVal OracleLoginData As ERPLogin) As String
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr, Flag As String
                Flag = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT008'"))
                Return Flag
            End Using
        Catch ex As Exception
            ErrorLogging("MaterialReplenishment-AllowEditDJ", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return "NO"
        End Try
    End Function

    Public Function AllowDiffDJ(ByVal OracleLoginData As ERPLogin) As String
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr, Flag As String
                Flag = da.ExecuteScalar(String.Format("select Value from T_Config where ConfigID = 'MMT009'"))
                Return Flag
            End Using
        Catch ex As Exception
            ErrorLogging("MaterialReplenishment-AllowDiffDJ", OracleLoginData.User, ex.Message & ex.Source, "E")
            Return "NO"
        End Try
    End Function

    Public Function DJCmp_GetBoxInfo(ByVal ID As String, ByVal OracleLoginData As ERPLogin) As dj_response
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr, GetDjInfo_Way As String
                Dim rs As New Object

                Dim CntSN As New DataSet
                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                    Sqlstr = String.Format(String.Format("SELECT OrgCode,COUNT(distinct ProductSerialNo) FROM T_Shippment WITH (nolock) where CartonID = '{0}' and statuscode = 1 group by OrgCode", ID))
                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                    Sqlstr = String.Format(String.Format("SELECT OrgCode, COUNT(distinct ProductSerialNo) FROM T_Shippment WITH (nolock) where PalletID = '{0}' and statuscode = 1 group by OrgCode", ID))
                End If
                CntSN = da.ExecuteDataSet(Sqlstr, "Cnt")
                If CntSN.Tables(0).Rows.Count > 1 OrElse (CntSN.Tables(0).Rows.Count = 1 AndAlso FixNull(CntSN.Tables(0).Rows(0)("OrgCode")) <> OracleLoginData.OrgCode) Then
                    DJCmp_GetBoxInfo.dsCLID = Nothing
                    DJCmp_GetBoxInfo.errormsg = "BoxID belong to another OrgCode"
                    DJCmp_GetBoxInfo.flag = "N"
                    Exit Function
                End If
                Sqlstr = String.Format(String.Format("select Value from T_Config where ConfigID = 'MMT010'"))
                GetDjInfo_Way = da.ExecuteScalar(Sqlstr)
                If GetDjInfo_Way Is Nothing OrElse GetDjInfo_Way = "" Then
                    DJCmp_GetBoxInfo.dsCLID = Nothing
                    DJCmp_GetBoxInfo.errormsg = "Setting of MMT010 in T_Config is not ready"
                    DJCmp_GetBoxInfo.flag = "N"
                    Exit Function

                ElseIf GetDjInfo_Way = "1" Then
                    If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                        rs = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where CLID = '{0}' and (statuscode = 1 or statuscode = 0 and lasttransaction <> 'DJ_Reversal')", ID))
                    ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                        rs = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where BoxID = '{0}' and (statuscode = 1 or statuscode = 0 and lasttransaction <> 'DJ_Reversal')", ID))
                    End If

                    If Not rs Is Nothing Then
                        DJCmp_GetBoxInfo.dsCLID = Nothing
                        DJCmp_GetBoxInfo.errormsg = "Invalid BoxID/PalletID"
                        DJCmp_GetBoxInfo.flag = "N"
                        Exit Function
                    Else
                        'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                        '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev, T_POQty.PO FROM T_Shippment WITH (nolock) INNER JOIN T_POQty WITH (nolock) ON T_Shippment.OrgCode = T_POQty.OrgCode AND T_Shippment.ProdOrder = T_POQty.PO WHERE (T_Shippment.CartonID = '{0}') AND (T_Shippment.OrgCode IS NULL OR T_Shippment.OrgCode = '{1}') AND (T_Shippment.StatusCode = '1') GROUP BY T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev,T_POQty.PO ", ID, OracleLoginData.OrgCode))
                        'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                        '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev, T_POQty.PO FROM T_Shippment WITH (nolock) INNER JOIN T_POQty WITH (nolock) ON T_Shippment.OrgCode = T_POQty.OrgCode AND T_Shippment.ProdOrder = T_POQty.PO WHERE (T_Shippment.PalletID = '{0}') and (T_Shippment.OrgCode is NULL or T_Shippment.OrgCode = '{1}') AND (T_Shippment.StatusCode = '1') GROUP BY T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev,T_POQty.PO ", ID, OracleLoginData.OrgCode))
                        'End If
                        If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                            Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev, T_POQty.PO FROM T_Shippment WITH (nolock) INNER JOIN T_POQty WITH (nolock) ON T_Shippment.ProdOrder = T_POQty.PO WHERE (T_Shippment.CartonID = '{0}') AND (T_Shippment.StatusCode = '1') GROUP BY T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev,T_POQty.PO ", ID))
                        ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                            Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev, T_POQty.PO FROM T_Shippment WITH (nolock) INNER JOIN T_POQty WITH (nolock) ON T_Shippment.ProdOrder = T_POQty.PO WHERE (T_Shippment.PalletID = '{0}') AND (T_Shippment.StatusCode = '1') GROUP BY T_Shippment.PalletID, T_Shippment.CartonID, T_POQty.Model, T_POQty.ModelRev,T_POQty.PO ", ID))
                        End If

                        DJCmp_GetBoxInfo.dsCLID = da.ExecuteDataSet(Sqlstr, "BoxInfo")
                        If Not DJCmp_GetBoxInfo.dsCLID Is Nothing AndAlso DJCmp_GetBoxInfo.dsCLID.Tables.Count > 0 AndAlso DJCmp_GetBoxInfo.dsCLID.Tables(0).Rows.Count > 0 Then
                            DJCmp_GetBoxInfo.errormsg = ""
                            DJCmp_GetBoxInfo.flag = "Y"
                            Exit Function
                        Else
                            DJCmp_GetBoxInfo.errormsg = "Invalid BoxID/PalletID"
                            DJCmp_GetBoxInfo.flag = "N"
                            Exit Function
                        End If
                    End If

                ElseIf GetDjInfo_Way = "2" Then
                    Dim DJList, DJInfo As DataSet
                    Dim i As Integer
                    Dim Model, ModelRev, Flag, Server_Flag As String

                    If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                        rs = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where CLID = '{0}' and (statuscode = 1 or statuscode = 0 and lasttransaction <> 'DJ_Reversal')", ID))
                    ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                        rs = da.ExecuteScalar(String.Format("select CLID BoxID from T_CLMaster with (nolock) where BoxID = '{0}' and (statuscode = 1 or statuscode = 0 and lasttransaction <> 'DJ_Reversal')", ID))
                    End If

                    If Not rs Is Nothing Then
                        DJCmp_GetBoxInfo.dsCLID = Nothing
                        DJCmp_GetBoxInfo.errormsg = "Invalid BoxID/PalletID"
                        DJCmp_GetBoxInfo.flag = "N"
                        Exit Function
                    Else
                        'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                        '    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM T_Shippment WITH (nolock) where CartonID = '{0}' and OrgCode = '{1}' and statuscode = 1", ID, OracleLoginData.OrgCode))
                        'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                        '    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM T_Shippment WITH (nolock) where PalletID = '{0}' and OrgCode = '{1}' and statuscode = 1", ID, OracleLoginData.OrgCode))
                        'End If
                        If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                            Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM T_Shippment WITH (nolock) where CartonID = '{0}' and statuscode = 1", ID))
                        ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                            Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM T_Shippment WITH (nolock) where PalletID = '{0}' and statuscode = 1", ID))
                        End If

                        DJList = New DataSet
                        DJList = da.ExecuteDataSet(Sqlstr, "DJ")

                        'check record in APP68
                        Flag = FixNull(GetServerFlag(OracleLoginData))
                        If Flag = "YES" Then
                            If DJList Is Nothing OrElse DJList.Tables.Count < 1 OrElse DJList.Tables(0).Rows.Count < 1 Then
                                Server_Flag = "APP68"
                                'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) where CartonID = '{0}' and OrgCode = '{1}' and statuscode = 1", ID, OracleLoginData.OrgCode))
                                'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) where PalletID = '{0}' and OrgCode = '{1}' and statuscode = 1", ID, OracleLoginData.OrgCode))
                                'End If
                                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) where CartonID = '{0}' and statuscode = 1", ID))
                                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT distinct ProdOrder FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) where PalletID = '{0}' and statuscode = 1", ID))
                                End If

                                DJList = New DataSet
                                DJList = da.ExecuteDataSet(Sqlstr, "DJ")
                            Else
                                Server_Flag = "APP58"
                            End If
                        Else
                            Server_Flag = ""
                        End If

                        If DJList Is Nothing OrElse DJList.Tables.Count < 1 OrElse DJList.Tables(0).Rows.Count < 1 Then
                            DJCmp_GetBoxInfo.dsCLID = Nothing
                            DJCmp_GetBoxInfo.errormsg = "Invalid BoxID/PalletID"
                            DJCmp_GetBoxInfo.flag = "N"
                            Exit Function
                        Else
                            For i = 0 To DJList.Tables(0).Rows.Count - 1
                                DJInfo = New DataSet
                                DJInfo = GetDJInfoFromERP(OracleLoginData, DJList.Tables(0).Rows(i)(0))
                                If Not DJInfo Is Nothing AndAlso DJInfo.Tables.Count > 0 AndAlso DJInfo.Tables(0).Rows.Count > 0 Then
                                    If i = 0 OrElse (Model = "" AndAlso ModelRev = "") Then
                                        Model = DJInfo.Tables(0).Rows(0)("product_number")
                                        ModelRev = DJInfo.Tables(0).Rows(0)("dj_revision")
                                    Else
                                        If Not (Model = DJInfo.Tables(0).Rows(0)("product_number") AndAlso ModelRev = DJInfo.Tables(0).Rows(0)("dj_revision")) Then
                                            DJCmp_GetBoxInfo.dsCLID = Nothing
                                            DJCmp_GetBoxInfo.errormsg = "DJ of BoxID/PalletID contain different Model/Rev"
                                            DJCmp_GetBoxInfo.flag = "N"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Next
                            If ModelRev = "" Then
                                DJCmp_GetBoxInfo.dsCLID = Nothing
                                DJCmp_GetBoxInfo.errormsg = "Can't get ModelRev from DJ in Oracle. Pls contact IT."
                                DJCmp_GetBoxInfo.flag = "N"
                                Exit Function
                            End If

                            If Server_Flag = "" OrElse Server_Flag = "APP58" Then
                                'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM T_Shippment WITH (nolock) WHERE CartonID = '{1}' AND (OrgCode IS NULL OR OrgCode = '{2}') AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID, OracleLoginData.OrgCode))
                                'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM T_Shippment WITH (nolock) WHERE PalletID = '{1}' AND (OrgCode is NULL or OrgCode = '{2}') AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID, OracleLoginData.OrgCode))
                                'End If
                                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM T_Shippment WITH (nolock) WHERE CartonID = '{1}' AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID))
                                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM T_Shippment WITH (nolock) WHERE PalletID = '{1}' AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID))
                                End If

                            ElseIf Server_Flag = "APP68" Then
                                'If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) WHERE CartonID = '{1}' AND (OrgCode IS NULL OR OrgCode = '{2}') AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID, OracleLoginData.OrgCode))
                                'ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                '    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) WHERE PalletID = '{1}' AND (OrgCode is NULL or OrgCode = '{2}') AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID, OracleLoginData.OrgCode))
                                'End If
                                If Mid(ID, 1, 1) <> "P" AndAlso Mid(ID, 1, 1) <> "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) WHERE CartonID = '{1}' AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID))
                                ElseIf Mid(ID, 1, 1) = "P" OrElse Mid(ID, 1, 1) = "1" Then
                                    Sqlstr = String.Format(String.Format("SELECT COUNT(distinct ProductSerialNo) AS COUNT, PalletID, CartonID, ProdOrder as PO, Model, '{0}' as ModelRev FROM ldetraceaddition.eTraceV2.dbo.T_Shippment WITH (nolock) WHERE PalletID = '{1}' AND StatusCode = '1' GROUP BY PalletID, CartonID, ProdOrder,Model", ModelRev, ID))
                                End If

                            End If

                            DJCmp_GetBoxInfo.dsCLID = da.ExecuteDataSet(Sqlstr, "BoxInfo")
                            If Not DJCmp_GetBoxInfo.dsCLID Is Nothing AndAlso DJCmp_GetBoxInfo.dsCLID.Tables.Count > 0 AndAlso DJCmp_GetBoxInfo.dsCLID.Tables(0).Rows.Count > 0 Then
                                DJCmp_GetBoxInfo.errormsg = ""
                                DJCmp_GetBoxInfo.flag = "Y"
                                Exit Function
                            Else
                                DJCmp_GetBoxInfo.errormsg = "Invalid BoxID/PalletID"
                                DJCmp_GetBoxInfo.flag = "N"
                                Exit Function
                            End If
                        End If
                    End If
                End If
                Return DJCmp_GetBoxInfo
            End Using
        Catch ex As Exception
            ErrorLogging("MaterialReplenishment-DJCmp_GetBoxInfo", OracleLoginData.User, ex.Message & ex.Source, "E")
            DJCmp_GetBoxInfo.dsCLID = Nothing
            DJCmp_GetBoxInfo.errormsg = ex.Message.ToString
            DJCmp_GetBoxInfo.flag = "N"
            Return DJCmp_GetBoxInfo
        End Try
    End Function

    Public Function CheckDJforCLID(ByVal CLID As String, ByVal DJ As String, ByVal Qty As String, ByVal OracleLoginData As ERPLogin) As String
        Dim DS As New DataSet
        Try
            Using da As DataAccess = GetDataAccess()
                Dim Sqlstr, Msg, Flag As String
                Dim i As Integer

                'If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                '    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' and (OrgCode is NULL or OrgCode = '{1}') group by ProdOrder,PalletID,StatusCode", CLID, OracleLoginData.OrgCode)
                'Else
                '    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' and (OrgCode is NULL or OrgCode = '{1}') group by ProdOrder,CartonID,StatusCode", CLID, OracleLoginData.OrgCode)
                'End If
                If (Mid(CLID, 3, 1) = "B" AndAlso Len(CLID) = 20) OrElse Mid(CLID, 1, 1) = "P" OrElse Mid(CLID, 1, 1) = "1" Then
                    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' group by ProdOrder,PalletID,StatusCode", CLID)
                Else
                    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' group by ProdOrder,CartonID,StatusCode", CLID)
                End If

                DS = New DataSet
                DS = da.ExecuteDataSet(Sqlstr, "DJ")

                'check record in APP68
                Flag = FixNull(GetServerFlag(OracleLoginData))
                If Flag = "YES" Then
                    If (Not DS Is Nothing AndAlso DS.Tables.Count > 0 AndAlso DS.Tables("DJ").Rows.Count > 0) Then

                    Else

                        'If (Mid(CLID, 3, 1) = "B" And Len(CLID) = 20) Or Mid(CLID, 1, 1) = "P" Then
                        '    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' and (OrgCode is NULL or OrgCode = '{1}') group by ProdOrder,PalletID,StatusCode", CLID, OracleLoginData.OrgCode)
                        'Else
                        '    Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' and (OrgCode is NULL or OrgCode = '{1}') group by ProdOrder,CartonID,StatusCode", CLID, OracleLoginData.OrgCode)
                        'End If
                        If (Mid(CLID, 3, 1) = "B" AndAlso Len(CLID) = 20) OrElse Mid(CLID, 1, 1) = "P" OrElse Mid(CLID, 1, 1) = "1" Then
                            Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where PalletID = '{0}' and StatusCode = '1' group by ProdOrder,PalletID,StatusCode", CLID)
                        Else
                            Sqlstr = String.Format("Select ProdOrder, count(*) as Qty from ldetraceaddition.eTraceV2.dbo.T_Shippment with (nolock) where CartonID = '{0}' and StatusCode = '1' group by ProdOrder,CartonID,StatusCode", CLID)
                        End If

                        DS = New DataSet
                        DS = da.ExecuteDataSet(Sqlstr, "DJ")
                    End If
                End If

                If DJ = "" And Qty = "" Then   'When check DJ information for boxid/palletid
                    If Not DS Is Nothing AndAlso DS.Tables.Count > 0 AndAlso DS.Tables("DJ").Rows.Count > 0 Then
                        If DS.Tables("DJ").Rows.Count > 1 Then
                            Msg = DS.Tables("DJ").Rows.Count & " DJ in BoxID/PalletID."
                            For i = 0 To DS.Tables("DJ").Rows.Count - 1
                                Msg = Msg & "DJ" & i + 1 & ":" & DS.Tables("DJ").Rows(i)("ProdOrder") & ", Qty: " & DS.Tables("DJ").Rows(i)("Qty") & ". "
                            Next
                            CheckDJforCLID = Msg
                        Else
                            CheckDJforCLID = ""
                        End If
                    Else
                        CheckDJforCLID = ""
                    End If
                Else              'When check DJ information within boxid/palletid
                    Dim DR() As DataRow = Nothing

                    DR = DS.Tables("DJ").Select("ProdOrder = '" & DJ & "'")
                    If DR.Length < 1 Then
                        CheckDJforCLID = "BoxID " & CLID & " not contain DJ " & DJ
                    Else
                        If DS.Tables("DJ").Rows.Count = 1 Then
                            CheckDJforCLID = CLID & " only contain DJ " & DJ & "! Please do DJ Completion per scan BoxID"
                        Else
                            If DR(0)("Qty") <> CDec(Qty) Then
                                CheckDJforCLID = "To complete Qty: " & Qty & " not equal to Qty: " & DR(0)("Qty") & " in Ref BoxID"
                            Else
                                CheckDJforCLID = ""
                            End If
                        End If
                    End If
                End If

            End Using
        Catch ex As Exception
            ErrorLogging("DJ_Completion-CheckDJforCLID", OracleLoginData.User, ex.Message & ex.Source, "E")
        End Try
    End Function

End Class
