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

Public Class Customs
    Inherits PublicFunction

    Public Function Get_item_wastage(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_item_wastage"
                oda.SelectCommand.Parameters.Add("p_org_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_item_wastage", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Insert_item_wastage(ByVal LoginData As ERPLogin, ByVal Itemnum As String, ByVal Wastage As Decimal, ByVal Description As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.insert_item_wastage"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                OC.Parameters.Add("p_item_num", OracleType.VarChar, 150).Value = Itemnum
                OC.Parameters.Add("p_wastage", OracleType.Number).Value = Wastage
                OC.Parameters.Add("p_description", OracleType.VarChar, 240).Value = Description
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                Else
                    flag = "E"
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Insert_item_wastage", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function delete_xxetr_item_wastage(ByVal RowID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.delete_xxetr_item_wastage"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 80).Value = RowID
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                Else
                    flag = "E"
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_xxetr_item_wastage", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Update_xxetr_item_wastage(ByVal RowID As String, ByVal wastage As Decimal, ByVal description As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.update_xxetr_item_wastage"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 60).Value = RowID
                OC.Parameters.Add("p_wastage", OracleType.VarChar, 60).Value = wastage
                OC.Parameters.Add("p_description", OracleType.VarChar, 60).Value = description
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Update_xxetr_item_wastage", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Get_Secondary_Inventory(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_secondary_inventory"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_Secondary_Inventory", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_Sub_Locator(ByVal LoginData As ERPLogin, ByVal subinventory_code As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_sub_locator"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_inventory_code", OracleType.VarChar).Value = subinventory_code
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_Sub_Locator", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_List_Sub_Locator(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_xxetr_second_inventory"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_List_Sub_Locator", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Find_List_Sub_Locator(ByVal LoginData As ERPLogin, ByVal Sub_type As String, ByVal Subinventory As String, ByVal Locator As String, ByVal Indecator As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.find_xxetr_second_inventory"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_sub_type", OracleType.VarChar).Value = Sub_type
                oda.SelectCommand.Parameters.Add("p_sub_inventory", OracleType.VarChar).Value = Subinventory
                oda.SelectCommand.Parameters.Add("p_sub_locator", OracleType.VarChar).Value = Locator
                oda.SelectCommand.Parameters.Add("p_indecator", OracleType.VarChar).Value = Indecator
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Find_List_Sub_Locator", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Add_Sub_Locator(ByVal LoginData As ERPLogin, ByVal SUBType As String, ByVal SUBINV As String, ByVal Locator As String, ByVal Indecator As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.insert_xxetr_second_inventory"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                OC.Parameters.Add("p_sub_type", OracleType.VarChar, 15).Value = SUBType
                OC.Parameters.Add("p_sub_inventory", OracleType.VarChar, 20).Value = SUBINV
                OC.Parameters.Add("p_sub_locator", OracleType.VarChar, 60).Value = Locator
                If Indecator = "" Then
                    OC.Parameters.Add("p_indecator", OracleType.VarChar, 10).Value = DBNull.Value
                Else
                    OC.Parameters.Add("p_indecator", OracleType.VarChar, 10).Value = Indecator
                End If

                OC.Parameters.Add("p_user", OracleType.VarChar, 50).Value = LoginData.User
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                Else
                    flag = "E"
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Add_Sub_Locator", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Delete_Sub_Locator(ByVal RowID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.delete_xxetr_second_inventory"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 80).Value = RowID
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                Else
                    flag = "E"
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_Sub_Locator", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Validate_item_number(ByVal LoginData As ERPLogin, ByVal ItemNo As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim Array_Item() As String
            Dim j As Integer
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.validate_item_number"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                'OC.Parameters.Add("p_item_number", OracleType.VarChar, 60).Value = ItemNo
                OC.Parameters.Add("p_item_number", OracleType.VarChar, 60)
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                Array_Item = Split(ItemNo, ",")
                For j = LBound(Array_Item) To UBound(Array_Item)
                    If Trim(Array_Item(j)) <> "" Then
                        OC.Parameters("p_item_number").Value = Trim(Array_Item(j))
                        result = CInt(OC.ExecuteNonQuery())
                        If OC.Parameters("x_flag").Value.ToString = "1" Then
                            flag = "S"
                        ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                            flag = OC.Parameters("x_msg").Value.ToString
                            OC.Connection.Close()
                            Exit Try
                        End If
                    End If
                Next
                OC.Connection.Close()
            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Validate_item_number", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function add_floor_stock_material(ByVal LoginData As ERPLogin, ByVal ItemNo As String, ByVal usage1 As String, ByVal usage2 As String, ByVal usage3 As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.insert_floor_stock_material"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                OC.Parameters.Add("p_item_number", OracleType.VarChar, 4000).Value = ItemNo
                OC.Parameters.Add("p_usage1", OracleType.VarChar, 20).Value = usage1
                OC.Parameters.Add("p_usage2", OracleType.VarChar, 20).Value = usage2
                OC.Parameters.Add("p_usage3", OracleType.VarChar, 20).Value = usage3
                OC.Parameters.Add("p_user", OracleType.VarChar, 60).Value = LoginData.User
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-add_floor_stock_material", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Get_List_Floor_Stock(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_floor_stock_material"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_List_Floor_Stock", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Find_List_Floor_Stock(ByVal LoginData As ERPLogin, ByVal Item As String, ByVal Usage1 As String, ByVal Usage2 As String, ByVal Usage3 As String) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.find_floor_stock_material"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("p_item_no", OracleType.VarChar)
                oda.SelectCommand.Parameters.Add("p_usage1", OracleType.VarChar)
                oda.SelectCommand.Parameters.Add("p_usage2", OracleType.VarChar)
                oda.SelectCommand.Parameters.Add("p_usage3", OracleType.VarChar)
                If Item = "" Then
                    oda.SelectCommand.Parameters("p_item_no").Value = DBNull.Value
                Else
                    oda.SelectCommand.Parameters("p_item_no").Value = Item
                End If
                If Usage1 = "" Then
                    oda.SelectCommand.Parameters("p_usage1").Value = DBNull.Value
                Else
                    oda.SelectCommand.Parameters("p_usage1").Value = Usage1
                End If
                If Usage2 = "" Then
                    oda.SelectCommand.Parameters("p_usage2").Value = DBNull.Value
                Else
                    oda.SelectCommand.Parameters("p_usage2").Value = Usage2
                End If
                If Usage3 = "" Then
                    oda.SelectCommand.Parameters("p_usage3").Value = DBNull.Value
                Else
                    oda.SelectCommand.Parameters("p_usage3").Value = Usage3
                End If
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Find_List_Floor_Stock", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Delete_Floor_Stock(ByVal RowID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.delete_floor_stock_material"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 60).Value = RowID
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_Floor_Stock", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Update_Floor_Stock(ByVal RowID As String, ByVal Usage1 As String, ByVal Usage2 As String, ByVal Usage3 As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.update_floor_stock_material"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 60).Value = RowID
                OC.Parameters.Add("p_usage1", OracleType.VarChar, 60).Value = Usage1
                OC.Parameters.Add("p_usage2", OracleType.VarChar, 60).Value = Usage2
                OC.Parameters.Add("p_usage3", OracleType.VarChar, 60).Value = Usage3
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Update_Floor_Stock", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Get_Generic_Disposition(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_generic_dispositions"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_Generic_Disposition", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function Get_List_Generic_Disposition(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_xxetr_generic_dispositions"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_List_Generic_Disposition", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function add_generic_dispositions(ByVal LoginData As ERPLogin, ByVal DispName As String, ByVal DispDESC As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.insert_generic_dispositions"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                OC.Parameters.Add("p_disp_name", OracleType.VarChar, 60).Value = DispName
                OC.Parameters.Add("p_disp_desc", OracleType.VarChar, 60).Value = DispDESC
                OC.Parameters.Add("p_user", OracleType.VarChar, 60).Value = LoginData.User
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-add_generic_dispositions", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Delete_Generic_Dispositions(ByVal RowID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.delete_xxetr_generic_dis"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 60).Value = RowID
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_Generic_Dispositions", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Get_List_Item_Mapping(ByVal LoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()

            Dim oda As OracleDataAdapter = da.Oda_Sele()
            Dim ds As New DataSet()
            Try
                oda.SelectCommand.CommandType = CommandType.StoredProcedure
                oda.SelectCommand.CommandText = "apps.xxetr_custom_data_upload.get_item_number_mapping"
                oda.SelectCommand.Parameters.Add("p_organization_code", OracleType.VarChar).Value = LoginData.OrgCode
                oda.SelectCommand.Parameters.Add("o_ref_cursor", OracleType.Cursor).Direction = ParameterDirection.Output
                oda.SelectCommand.Connection.Open()
                oda.Fill(ds)

                oda.SelectCommand.Connection.Close()
                Return ds

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Get_List_Item_Mapping", LoginData.User.ToUpper, oe.Message)
                Return ds
            Finally
                If oda.SelectCommand.Connection.State <> ConnectionState.Closed Then oda.SelectCommand.Connection.Close()
            End Try
        End Using
    End Function

    Public Function add_List_Item_Mapping(ByVal LoginData As ERPLogin, ByVal Ora_segment1 As String, ByVal Iems_segment1 As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.insert_item_number_mapping"
                OC.Parameters.Add("p_org_code", OracleType.VarChar, 3).Value = LoginData.OrgCode
                OC.Parameters.Add("p_segment1", OracleType.VarChar, 60).Value = Ora_segment1
                OC.Parameters.Add("p_iems_segment1", OracleType.VarChar, 60).Value = Iems_segment1
                OC.Parameters.Add("p_user", OracleType.VarChar, 60).Value = LoginData.User
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-add_List_Item_Mapping", LoginData.User.ToUpper, oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function Delete_List_Item_Mapping(ByVal RowID As String) As String
        Using da As DataAccess = GetDataAccess()

            Dim result As Integer
            Dim flag, msg As String
            Dim OC As OracleCommand = da.OraCommand()

            Try
                OC.CommandType = CommandType.StoredProcedure
                OC.CommandText = "apps.xxetr_custom_data_upload.delete_item_number_mapping"
                OC.Parameters.Add("p_row_id", OracleType.VarChar, 60).Value = RowID
                OC.Parameters.Add("x_flag", OracleType.VarChar, 1).Direction = ParameterDirection.Output
                OC.Parameters.Add("x_msg", OracleType.VarChar, 500).Direction = ParameterDirection.Output

                OC.Connection.Open()
                result = CInt(OC.ExecuteNonQuery())
                OC.Connection.Close()

                If OC.Parameters("x_flag").Value.ToString = "1" Then
                    flag = "S"
                ElseIf OC.Parameters("x_flag").Value.ToString = "0" Then
                    flag = OC.Parameters("x_msg").Value.ToString
                End If

            Catch oe As Exception
                ErrorLogging("LabelBasicFunction-Delete_List_Item_Mapping", "", oe.Message & oe.Source, "E")
                flag = "E"
            Finally
                If OC.Connection.State <> ConnectionState.Closed Then OC.Connection.Close()
            End Try
            Return flag
        End Using
    End Function

    Public Function LoadCustomReportData(ByVal dataType As String, ByVal pannelVisible As Boolean, ByVal checkBox As String, ByVal selectedTable As String, ByVal erpLogin As ERPLogin, ByVal dsInput As DataSet) As DataSet
        Dim dsError As DataSet = New DataSet
        Dim tableError As DataTable = New DataTable("error")
        Dim dataColumn As DataColumn
        dataColumn = New Data.DataColumn("Data Type", System.Type.GetType("System.String"))
        tableError.Columns.Add(dataColumn)
        dataColumn = New Data.DataColumn("Identified Column", System.Type.GetType("System.String"))
        tableError.Columns.Add(dataColumn)
        dataColumn = New Data.DataColumn("Error Message", System.Type.GetType("System.String"))
        tableError.Columns.Add(dataColumn)

        Using da As DataAccess = GetDataAccess()
            Dim oCommand As OracleCommand = da.OraCommand()
            Try
                If oCommand.Connection.State = ConnectionState.Closed Then
                    oCommand.Connection.Open()
                End If
                oCommand.CommandType = CommandType.StoredProcedure
                If dataType <> "ADDITIONAL DATA" And pannelVisible = False Then
                    If dataType = "ALIAS ACCOUNT" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_generic_dispositions"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_disposition_name", OracleType.VarChar, 150)
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 2).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_disposition_name").Value = dr(0).ToString
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next
                    ElseIf dataType = "IEMS INVENTORY" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_iems_inventory"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_iems_segment1", OracleType.VarChar, 150)
                        oCommand.Parameters.Add("p_iems_desc", OracleType.VarChar, 200)
                        oCommand.Parameters.Add("p_iems_uom", OracleType.VarChar, 50)
                        oCommand.Parameters.Add("p_iems_price", OracleType.Number)
                        oCommand.Parameters.Add("p_iems_onhand", OracleType.Number)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 60).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        If CountNotNULLDataRow(dsInput.Tables(dataType)) = 0 Then
                            DeleteInputData(erpLogin.OrgCode, dataType)
                        Else
                            For Each dr As DataRow In dsInput.Tables(dataType).Rows
                                If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                    oCommand.Parameters.Item("p_iems_segment1").Value = dr(0).ToString.Trim
                                    oCommand.Parameters.Item("p_iems_desc").Value = dr(1).ToString.Trim
                                    oCommand.Parameters.Item("p_iems_uom").Value = dr(2).ToString.Trim
                                    If IsNumeric(dr(3)) And IsNumeric(dr(4)) Then
                                        oCommand.Parameters.Item("p_iems_price").Value = CDec(dr(3))
                                        oCommand.Parameters.Item("p_iems_onhand").Value = CDec(dr(4))
                                    Else
                                        Dim dRow As DataRow = tableError.NewRow
                                        dRow(0) = dataType
                                        dRow(1) = dr(0).ToString
                                        dRow(2) = "Column value should be number"
                                        tableError.Rows.Add(dRow)
                                        GoTo nextFor
                                    End If
                                    oCommand.ExecuteNonQuery()
                                    If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                        Dim dRow As DataRow = tableError.NewRow
                                        dRow(0) = dataType
                                        dRow(1) = dr(0).ToString
                                        dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                        tableError.Rows.Add(dRow)
                                    End If
                                    oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                End If
nextFor:
                            Next
                        End If
                    ElseIf dataType = "IEMS MAPPING" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_oracle_iems_mapping"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_ora_segment", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_iems_segment", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_ora_uom", OracleType.VarChar, 25)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 60).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        If CountNotNULLDataRow(dsInput.Tables(dataType)) = 0 Then
                            DeleteInputData(erpLogin.OrgCode, dataType)
                        Else
                            For Each dr As DataRow In dsInput.Tables(dataType).Rows
                                If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                    oCommand.Parameters.Item("p_ora_segment").Value = dr(0).ToString.Trim
                                    oCommand.Parameters.Item("p_iems_segment").Value = dr(1).ToString.Trim
                                    oCommand.Parameters.Item("p_ora_uom").Value = dr(2).ToString.Trim
                                    oCommand.ExecuteNonQuery()
                                    If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                        Dim dRow As DataRow = tableError.NewRow
                                        dRow(0) = dataType
                                        dRow(1) = dr(0).ToString
                                        dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                        tableError.Rows.Add(dRow)
                                    End If
                                    oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                End If
                            Next
                        End If
                    ElseIf dataType = "FLOOR MATERIAL" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_floor_stock_material"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_usage1", OracleType.Number)
                        oCommand.Parameters.Add("p_usage2", OracleType.Number)
                        oCommand.Parameters.Add("p_usage3", OracleType.Number)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 60).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString
                            If IsNumeric(dr(1)) And IsNumeric(dr(2)) And IsNumeric(dr(3)) Then
                                oCommand.Parameters.Item("p_usage1").Value = CDec(dr(1))
                                oCommand.Parameters.Item("p_usage2").Value = CDec(dr(2))
                                oCommand.Parameters.Item("p_usage3").Value = CDec(dr(3))
                            Else
                                Dim dRow As DataRow = tableError.NewRow
                                dRow(0) = dataType
                                dRow(1) = dr(0).ToString
                                dRow(2) = "Column value should be number"
                                tableError.Rows.Add(dRow)
                                GoTo nextFor1
                            End If
                            oCommand.ExecuteNonQuery()
                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                Dim dRow As DataRow = tableError.NewRow
                                dRow(0) = dataType
                                dRow(1) = dr(0).ToString
                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                tableError.Rows.Add(dRow)
                            End If
                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
nextFor1:
                        Next
                        '''''''''''''''''''''''''''''''''''''''''''
                    ElseIf dataType = "ITEM MASTER" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "AML E_PART" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "AML M_PART" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "ITEM W_USED" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "ORACLE BOM" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "IEMS BOM" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "ORACLE INVENTORY" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                    ElseIf dataType = "RECEIVING AND RTV" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.insert_batch_item_number"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_batch_type", OracleType.VarChar, 50).Value = dataType
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 1).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For Each dr As DataRow In dsInput.Tables(dataType).Rows
                            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                oCommand.ExecuteNonQuery()
                                If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                    Dim dRow As DataRow = tableError.NewRow
                                    dRow(0) = dataType
                                    dRow(1) = dr(0).ToString
                                    dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                    tableError.Rows.Add(dRow)
                                End If
                                oCommand.Parameters.Item("p_delete_flag").Value = "N"
                            End If
                        Next

                        '''''''''''''''''''''''''''''
                    End If
                ElseIf dataType = "ADDITIONAL DATA" And pannelVisible = True Then
                    If checkBox = "Y" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_input_onhand"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_input_type", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_onhand", OracleType.Number)
                        oCommand.Parameters.Add("p_uom", OracleType.VarChar, 20)
                        oCommand.Parameters.Add("p_iems_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        For i As Integer = 0 To dsInput.Tables.Count - 1
                            oCommand.Parameters.Item("p_delete_flag").Value = "Y"
                            Dim tbName As String = dsInput.Tables(i).TableName
                            If dsInput.Tables(i).TableName = "'DPT FG$'" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "DPT_FG")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "DPT_FG"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor2
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor2:
                                    Next
                                End If
                            ElseIf dsInput.Tables(i).TableName = "'DPT RAW$'" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "DPT_RAW")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "DPT_RAW"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor3
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor3:
                                    Next
                                End If
                            ElseIf dsInput.Tables(i).TableName = "'IN TRANSIT$'" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "IN_TRANSIT")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "IN_TRANSIT"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor4
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor4:
                                    Next
                                End If
                            ElseIf dsInput.Tables(i).TableName = "'NO GR$'" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "NO_GR")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "NO_GR"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor5
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor5:
                                    Next
                                End If
                            ElseIf dsInput.Tables(i).TableName = "'GR DIS$'" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "GR_DIS")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "GR_DIS"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor6
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Right(Left(tbName, tbName.Length - 2), Left(tbName, tbName.Length - 2).Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor6:
                                    Next
                                End If
                            ElseIf dsInput.Tables(i).TableName = "INDIRECT$" Then
                                If CountNotNULLDataRow(dsInput.Tables(dsInput.Tables(i).TableName)) = 0 Then
                                    DeleteInputData(erpLogin.OrgCode, "INDIRECT")
                                Else
                                    oCommand.Parameters.Item("p_input_type").Value = "INDIRECT"
                                    For Each dr As DataRow In dsInput.Tables(dsInput.Tables(i).TableName).Rows
                                        If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                            oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                            If IsNumeric(dr(1)) Then
                                                oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                            Else
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Left(tbName, tbName.Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = "Column value should be number"
                                                tableError.Rows.Add(dRow)
                                                GoTo nextFor7
                                            End If
                                            If dr(2) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_uom").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                            End If
                                            If dr(3) Is DBNull.Value Then
                                                oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                            Else
                                                oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                            End If
                                            oCommand.ExecuteNonQuery()
                                            If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                                Dim dRow As DataRow = tableError.NewRow
                                                dRow(0) = Left(tbName, tbName.Length - 1)
                                                dRow(1) = dr(0).ToString
                                                dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                                tableError.Rows.Add(dRow)
                                            End If
                                            oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                        End If
nextFor7:
                                    Next
                                End If
                            End If
                        Next
                    ElseIf checkBox = "N" Then
                        oCommand.CommandText = "apps.xxetr_custom_data_upload.load_input_onhand"
                        oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = erpLogin.OrgCode
                        oCommand.Parameters.Add("p_input_type", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_onhand", OracleType.Number)
                        oCommand.Parameters.Add("p_uom", OracleType.VarChar, 20)
                        oCommand.Parameters.Add("p_iems_segment1", OracleType.VarChar, 60)
                        oCommand.Parameters.Add("p_user_name", OracleType.VarChar, 60).Value = erpLogin.User.ToUpper
                        oCommand.Parameters.Add("p_delete_flag", OracleType.VarChar, 60).Value = "Y"
                        oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                        oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 200)
                        oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                        oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                        If selectedTable = "DPT FG" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "DPT_FG")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "DPT_FG"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor8
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor8:
                                Next
                            End If
                        ElseIf selectedTable = "DPT RAW" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "DPT_RAW")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "DPT_RAW"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor9
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor9:
                                Next
                            End If
                        ElseIf selectedTable = "IN TRANSIT" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "IN_TRANSIT")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "IN_TRANSIT"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor10
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor10:
                                Next
                            End If
                        ElseIf selectedTable = "GR DIS" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "GR_DIS")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "GR_DIS"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor11
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor11:
                                Next
                            End If
                        ElseIf selectedTable = "NO GR" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "NO_GR")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "NO_GR"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor12
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor12:
                                Next
                            End If
                        ElseIf selectedTable = "INDIRECT" Then
                            If CountNotNULLDataRow(dsInput.Tables(selectedTable)) = 0 Then
                                DeleteInputData(erpLogin.OrgCode, "INDIRECT")
                            Else
                                For Each dr As DataRow In dsInput.Tables(selectedTable).Rows
                                    If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                                        oCommand.Parameters.Item("p_input_type").Value = "INDIRECT"
                                        oCommand.Parameters.Item("p_segment1").Value = dr(0).ToString.Trim
                                        If IsNumeric(dr(1)) Then
                                            oCommand.Parameters.Item("p_onhand").Value = CDec(dr(1))
                                        Else
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = "Column value should be number"
                                            tableError.Rows.Add(dRow)
                                            GoTo nextFor13
                                        End If
                                        If dr(2) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_uom").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_uom").Value = dr(2).ToString
                                        End If
                                        If dr(3) Is DBNull.Value Then
                                            oCommand.Parameters.Item("p_iems_segment1").Value = ""
                                        Else
                                            oCommand.Parameters.Item("p_iems_segment1").Value = dr(3).ToString
                                        End If
                                        oCommand.ExecuteNonQuery()
                                        If oCommand.Parameters("o_error_code").Value.ToString = "0" Then
                                            Dim dRow As DataRow = tableError.NewRow
                                            dRow(0) = selectedTable
                                            dRow(1) = dr(0).ToString
                                            dRow(2) = oCommand.Parameters("o_error_message").Value.ToString
                                            tableError.Rows.Add(dRow)
                                        End If
                                        oCommand.Parameters.Item("p_delete_flag").Value = "N"
                                    End If
nextFor13:
                                Next
                            End If
                        End If
                    End If
                End If

            Catch ex As Exception
                Dim dRow As DataRow = tableError.NewRow
                dRow(0) = "Database"
                dRow(1) = "Exception"
                dRow(2) = ex.Message & ex.Source
                tableError.Rows.Add(dRow)
                ErrorLogging("LoadCustomReportData" & erpLogin.OrgCode, erpLogin.User, ex.Message & ex.Source, "E")
            Finally
                If oCommand.Connection.State = ConnectionState.Open Then
                    oCommand.Connection.Close()
                End If
                dsError.Tables.Add(tableError)
            End Try
            Return dsError
        End Using
    End Function

    Public Function DeleteInputData(ByVal organization_code As String, ByVal input_type As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim result_status As String = ""
            Dim oCommand As OracleCommand = da.OraCommand()
            Try
                If oCommand.Connection.State = ConnectionState.Closed Then
                    oCommand.Connection.Open()
                End If
                oCommand.CommandType = CommandType.StoredProcedure
                oCommand.CommandText = "apps.xxetr_custom_data_upload.delete_input_onhand"
                oCommand.Parameters.Add("p_organization_code", OracleType.VarChar, 60).Value = organization_code
                oCommand.Parameters.Add("p_input_type", OracleType.VarChar, 60).Value = input_type
                oCommand.Parameters.Add("o_error_code", OracleType.VarChar, 2)
                oCommand.Parameters.Add("o_error_message", OracleType.VarChar, 250)
                oCommand.Parameters("o_error_code").Direction = ParameterDirection.Output
                oCommand.Parameters("o_error_message").Direction = ParameterDirection.Output
                oCommand.ExecuteNonQuery()
                If oCommand.Parameters("o_error_code").Value = "1" Then
                    result_status = "1"
                ElseIf oCommand.Parameters("o_error_code").Value = "0" Then
                    result_status = "0"
                End If
            Catch ex As Exception
                ErrorLogging("DeleteInputData", "", ex.Message & ex.Source, "E")
            Finally
                If oCommand.Connection.State = ConnectionState.Open Then
                    oCommand.Connection.Close()
                End If
            End Try
            Return result_status
        End Using
    End Function

    Public Function CountNotNULLDataRow(ByVal dataTable As DataTable) As Integer
        Dim result_count As Integer = 0
        For Each dr As DataRow In dataTable.Rows
            If (Not String.IsNullOrEmpty(dr(0).ToString)) = True Then
                result_count = result_count + 1
            End If
        Next
        Return result_count
    End Function

End Class
