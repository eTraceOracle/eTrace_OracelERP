Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Security.Cryptography


Public Class Andon
    Inherits PublicFunction


    Public Function AndonInput(ByVal Model As String, ByVal Line As String, ByVal Station As String, ByVal status As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonInput '{0}', '{1}', '{2}','{3}'", Model, Line, Station, status)
                AndonInput = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonInput = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-SaveCartonID", "", ex.Message)
            End Try
        End Using

    End Function

    Public Function ExistsModel(ByVal Model As String) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Dim i As Integer
            Try
                strSQL = String.Format("select count(*) from T_ProductMaster with (nolock) where Model='{0}'", Model)
                i = Convert.ToString(da.ExecuteScalar(strSQL))
                If i = 0 Then
                    strSQL = String.Format("select count(*) from T_ProductCPN with (nolock) where Model='{0}'", Model)
                    i = Convert.ToString(da.ExecuteScalar(strSQL))
                End If
                Return i
            Catch ex As Exception
                Return 0
                ErrorLogging("Andon-ExistsModel", "", ex.Message)
            End Try
        End Using
    End Function

    Public Function AndonActualQty(ByVal Model As String, ByVal Line As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonActualQty '{0}', '{1}'", Model, Line)
                AndonActualQty = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonActualQty = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonActualQty", "", ex.Message)
            End Try
        End Using

    End Function

    Public Function AndonActualQtyofProcess(ByVal Model As String, ByVal Line As String, ByVal process As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonActualQtyofProcess '{0}', '{1}','{2}'", Model, Line, process)
                AndonActualQtyofProcess = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonActualQtyofProcess = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonActualQtyofProcess", "", ex.Message)
            End Try
        End Using

    End Function

    Public Function AndonFailedPercent(ByVal Model As String, ByVal Line As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonFailedPercent '{0}', '{1}'", Model, Line)
                AndonFailedPercent = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonFailedPercent = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonFailedPercent", "", ex.Message)
            End Try
        End Using
    End Function

    Public Function AndonFailedQty(ByVal Model As String, ByVal Line As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonFailedQty '{0}', '{1}'", Model, Line)
                AndonFailedQty = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonFailedQty = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonFailedQty", "", ex.Message)
            End Try
        End Using

    End Function

    Public Function AndonLineStopTime(ByVal Model As String, ByVal Line As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonLineStopTime '{0}', '{1}'", Model, Line)
                AndonLineStopTime = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonLineStopTime = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonLineStopTime", "", ex.Message)
            End Try
        End Using
    End Function

    Public Function AndonLineStopFreq(ByVal Model As String, ByVal Line As String) As String
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonLineStopFreq '{0}', '{1}'", Model, Line)
                AndonLineStopFreq = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonLineStopFreq = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("Andon-AndonLineStopFreq", "", ex.Message)
            End Try
        End Using
    End Function


    Public Function GetAndonProdSche(ByVal Line As String) As DataSet
        Dim i As Integer
        Dim dsAndonProdSche As DataSet = New DataSet("dsAndonProdSche")
        Dim dtAndonProdSche As DataTable = New DataTable("dtAndonProdSche")
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceDBConnString"))
        Try
            Dim da As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(String.Format("select StartTime, EndTime from T_AndonProdSche with (nolock) where Line='{0}' order by StartTime", Line), myConn)
            da.Fill(dtAndonProdSche)
            dsAndonProdSche.Tables.Add(dtAndonProdSche)

        Catch ex As Exception
            ErrorLogging("TDC-GetAndonProdSche", "", ex.Message)
        End Try
        Return dsAndonProdSche
    End Function


    Public Function UpdateAndonProdSche(ByVal ds As DataSet, ByVal Line As String) As String
        Dim cmdSQL As SqlClient.SqlCommand
        Dim myConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("eTraceDBConnString"))
        Dim strSQL As String
        Try
            myConn.Open()
            strSQL = String.Format("delete T_AndonProdSche where Line='{0}'", Line)
            cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
            cmdSQL.ExecuteNonQuery()
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                strSQL = String.Format("exec sp_AndonProdScheInsert '{0}', '{1}', '{2}'", Line, ds.Tables(0).Rows(i)("StartTime"), ds.Tables(0).Rows(i)("EndTime"))
                cmdSQL = New SqlClient.SqlCommand(strSQL, myConn)
                cmdSQL.ExecuteNonQuery()
            Next
            myConn.Close()
            UpdateAndonProdSche = ""
        Catch ex As Exception
            UpdateAndonProdSche = "Error: " & ex.Message & " Please contact Admin"
            ErrorLogging("TDC-UpdateAndonProdSche", "", ex.Message)
        Finally
            If myConn.State = ConnectionState.Open Then
                myConn.Close()
            End If
        End Try
    End Function


    Public Function AndonProjectedQty(ByVal Line As String, ByVal Target As Integer) As Integer
        Using da As DataAccess = GetDataAccess()
            Dim strSQL As String
            Try
                strSQL = String.Format("exec sp_AndonProjectedQty '{0}', {1}", Line, Target)
                AndonProjectedQty = Convert.ToString(da.ExecuteScalar(strSQL))
            Catch ex As Exception
                AndonProjectedQty = "Error: " & ex.Message & " Please contact Admin"
                ErrorLogging("TDC-AndonProjectedQty", "", ex.Message)
            End Try
        End Using

    End Function

End Class
