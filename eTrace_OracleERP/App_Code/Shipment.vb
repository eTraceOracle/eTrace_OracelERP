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


Public Structure GetShipment
    Dim Flag As String
    Dim Message As String
    Dim ds As DataSet
End Structure

Public Class Shipment
    Inherits PublicFunction

    Public Function GetModelRevision(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try

            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getModelAndRev", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5
            sda.SelectCommand.Parameters.Add("@ID", SqlDbType.VarChar, 30)
            sda.SelectCommand.Parameters("@ID").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@ID").Value = PalletID

            sda.SelectCommand.Parameters.Add("@Org", SqlDbType.VarChar, 30)
            sda.SelectCommand.Parameters("@Org").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@Org").Value = OracleLoginData.OrgCode

            If sda.SelectCommand.ExecuteScalar().ToString = "99" Then
            Else
                sda.Fill(dsmodel, "model")
            End If
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetModelRevision", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function GetModelRevDJ_WithType(ByVal PalletID As String, ByVal Type As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Using da As DataAccess = GetDataAccess()
            Dim ds As New DataSet
            Dim sqlstr, CLIDStatus, BoxIDStatus As String
            Try
                'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, "1:" & Type, "I")

                If Type = "DJ required" Then
                    If Mid(PalletID, 1, 1) <> "P" Then
                        sqlstr = String.Format("select top 1 StatusCode from T_CLMaster where CLID='{0}' and OrgCode='{1}'", PalletID, OracleLoginData.OrgCode)
                        CLIDStatus = da.ExecuteScalar(sqlstr)
                    Else
                        sqlstr = String.Format("select top 1 StatusCode from T_CLMaster where BoxID='{0}' and OrgCode='{1}'", PalletID, OracleLoginData.OrgCode)
                        CLIDStatus = da.ExecuteScalar(sqlstr)
                    End If

                    If CLIDStatus Is Nothing OrElse FixNull(CLIDStatus) = "" OrElse FixNull(CLIDStatus) = "0" Then
                        ds = GetModelRevDJ(PalletID, OracleLoginData)
                    Else
                        Exit Function
                    End If
                Else
                    'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, PalletID, "I")

                    If Mid(PalletID, 1, 1) <> "P" Then
                        'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, "2:" & Type, "I")

                        sqlstr = String.Format("select top 1 StatusCode from T_CLMaster where CLID='{0}' and OrgCode='{1}' and StatusCode = 0", PalletID, OracleLoginData.OrgCode)
                        CLIDStatus = da.ExecuteScalar(sqlstr)
                        sqlstr = String.Format("select top 1 StatusCode from T_Shippment where CartonID='{0}'  and (OrgCode is NULL or OrgCode='{1}') and StatusCode = 1", PalletID, OracleLoginData.OrgCode)
                        BoxIDStatus = da.ExecuteScalar(sqlstr)
                        'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, FixNull(CLIDStatus) & " & " & FixNull(BoxIDStatus), "I")

                        If (Not CLIDStatus Is Nothing AndAlso FixNull(CLIDStatus) <> "") AndAlso (Not BoxIDStatus Is Nothing AndAlso FixNull(BoxIDStatus) <> "") Then
                            ds = GetModelRevDJ(PalletID, OracleLoginData)
                        Else
                            Exit Function
                        End If
                    Else
                        'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, "3:" & Type, "I")

                        sqlstr = String.Format("select top 1 StatusCode from T_CLMaster where BoxID='{0}' and OrgCode='{1}' and StatusCode = 0", PalletID, OracleLoginData.OrgCode)
                        CLIDStatus = da.ExecuteScalar(sqlstr)
                        sqlstr = String.Format("select top 1 StatusCode from T_Shippment where PalletID='{0}'  and (OrgCode is NULL or OrgCode='{1}') and StatusCode = 1", PalletID, OracleLoginData.OrgCode)
                        BoxIDStatus = da.ExecuteScalar(sqlstr)
                        'ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, FixNull(CLIDStatus) & " & " & FixNull(BoxIDStatus), "I")

                        If (Not CLIDStatus Is Nothing AndAlso FixNull(CLIDStatus) <> "") AndAlso (Not BoxIDStatus Is Nothing AndAlso FixNull(BoxIDStatus) <> "") Then
                            ds = GetModelRevDJ(PalletID, OracleLoginData)
                        Else
                            Exit Function
                        End If
                    End If
                End If
                Return ds
            Catch ex As Exception
                ErrorLogging("GetModelRevDJ_WithType", OracleLoginData.User, ex.Message & ex.Source, "E")
            End Try
        End Using
    End Function

    Public Function GetModelRevDJ(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim dsmodel As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try

            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getModelRevDJ", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5
            sda.SelectCommand.Parameters.Add("@ID", SqlDbType.VarChar, 30)
            sda.SelectCommand.Parameters("@ID").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@ID").Value = PalletID

            sda.SelectCommand.Parameters.Add("@Org", SqlDbType.VarChar, 30)
            sda.SelectCommand.Parameters("@Org").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@Org").Value = OracleLoginData.OrgCode

            If sda.SelectCommand.ExecuteScalar().ToString = "99" Then
            Else
                sda.Fill(dsmodel, "model")
            End If
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetModelRevDJ", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return dsmodel
    End Function

    Public Function GetSNList(ByVal PalletID As String, ByVal OracleLoginData As ERPLogin) As DataSet
        Dim snlist As DataSet = New DataSet
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try

            myConn.Open()
            Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("sp_getSNList", myConn)
            sda.SelectCommand.CommandType = CommandType.StoredProcedure
            sda.SelectCommand.CommandTimeout = TimeOut_M5
            sda.SelectCommand.Parameters.Add("@ID", SqlDbType.VarChar, 30)
            sda.SelectCommand.Parameters("@ID").Direction = ParameterDirection.Input
            sda.SelectCommand.Parameters("@ID").Value = PalletID
            sda.Fill(snlist, "snlist")

            myConn.Close()
        Catch ex As Exception
            ErrorLogging("GetSNList", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return snlist
    End Function

    Public Function ChangeModel(ByVal BoxID() As String, ByVal NewModel As String, ByVal NewRev As String, ByVal OracleLoginData As ERPLogin) As String
        Dim result As String
        Dim i As Integer
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_ChangeModel", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@Model", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@Model").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Model").Value = NewModel

            cmdSQL.Parameters.Add("@Rev", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@Rev").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Rev").Value = NewRev

            cmdSQL.Parameters.Add("@user", SqlDbType.VarChar, 20)
            cmdSQL.Parameters("@user").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@user").Value = OracleLoginData.User

            cmdSQL.Parameters.Add("@ID", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@ID").Direction = ParameterDirection.Input

            For i = 0 To BoxID.Length - 1
                cmdSQL.Parameters("@ID").Value = BoxID(i)
                result = Convert.ToString(cmdSQL.ExecuteScalar())
            Next
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("ChangeModel", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return result
    End Function

    Public Function Change_Model(ByVal BoxID() As String, ByVal NewModel As String, ByVal NewRev As String, ByVal NewDJ As String, ByVal OracleLoginData As ERPLogin) As String
        Dim result As String
        Dim i As Integer
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceDBConnString"))
        Try
            myConn.Open()
            cmdSQL = New SqlClient.SqlCommand("sp_Change_Model", myConn)
            cmdSQL.CommandType = CommandType.StoredProcedure
            cmdSQL.CommandTimeout = TimeOut_M5

            cmdSQL.Parameters.Add("@Model", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@Model").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Model").Value = NewModel

            cmdSQL.Parameters.Add("@Rev", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@Rev").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@Rev").Value = NewRev

            cmdSQL.Parameters.Add("@DJ", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@DJ").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@DJ").Value = NewDJ

            cmdSQL.Parameters.Add("@user", SqlDbType.VarChar, 20)
            cmdSQL.Parameters("@user").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@user").Value = OracleLoginData.User

            cmdSQL.Parameters.Add("@OrgCode", SqlDbType.VarChar, 20)
            cmdSQL.Parameters("@OrgCode").Direction = ParameterDirection.Input
            cmdSQL.Parameters("@OrgCode").Value = OracleLoginData.OrgCode

            cmdSQL.Parameters.Add("@ID", SqlDbType.VarChar, 50)
            cmdSQL.Parameters("@ID").Direction = ParameterDirection.Input

            For i = 0 To BoxID.Length - 1
                cmdSQL.Parameters("@ID").Value = BoxID(i)
                result = Convert.ToString(cmdSQL.ExecuteScalar())
            Next
            myConn.Close()
        Catch ex As Exception
            ErrorLogging("Change_Model", OracleLoginData.User, ex.Message & ex.Source, "E")
        Finally
            If myConn.State <> ConnectionState.Closed Then myConn.Close()
        End Try
        Return result
    End Function
End Class
