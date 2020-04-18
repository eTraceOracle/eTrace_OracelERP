
Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient


Public Class OneToOne
    Inherits PublicFunction

    Public otoConn As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceOTOConnString"))

#Region "Component Verify"

    Public Function IsAuthorizedPCB(ByVal model As String, ByVal pcb As String, ByVal userName As String) As String
        Using da As DataAccess = GetDataAccess()
            Try
                Dim sql As String = String.Format("exec SP_IsAuthorizedPCB '{0}','{1}'", model, pcb)
                Dim msg = da.ExecuteScalar(sql, otoConn).ToString()
                Return msg

            Catch ex As Exception
                ErrorLoggingOTO("OTO-IsAuthorizedPCB", userName, ex.Message & ex.Source)
                Return Nothing
            End Try
        End Using
    End Function

#End Region

#Region "Kanban Management"

    ''Public Function GetKBLabelInfo(ByVal DJ As String, ByVal CircuitCode As String, ByVal PN As String) As DataSet

    ''End Function

    ''Public Function IsHeatThing(ByVal userName As String, ByVal kbCode As String) As Boolean
    ''    Using da As DataAccess = GetDataAccess()
    ''        Try
    ''            Dim sql As String = String.Format("exec SP_IsHeatThing '{0}'", kbCode)
    ''            ''ErrorLoggingOTO("OTO-IsHeatThing", userName, "test")
    ''            Dim msg As String = da.ExecuteScalar(sql, otoConn)
    ''            ''ErrorLoggingOTO("OTO-IsHeatThing", userName, msg)
    ''            If (String.Equals(msg, "SUBA", StringComparison.OrdinalIgnoreCase)) Then
    ''                Return True
    ''            End If
    ''        Catch ex As Exception
    ''            ErrorLoggingOTO("OTO-IsHeatThing", userName, ex.Message & ex.Source & " " & kbCode)
    ''            Return False
    ''        End Try
    ''        Return False
    ''    End Using
    ''End Function

    ''Public Function GetPTImplementDJInfo(ByVal userName As String, ByVal dj As String) As DataTable
    ''    Using da As DataAccess = GetDataAccess()
    ''        Try
    ''            Dim sql As String = String.Format("exec SP_GetTImplementDJInfo '{0}'", dj)
    ''            Return da.ExecuteDataTable2(sql, CommandType.Text, otoConn)
    ''        Catch ex As Exception
    ''            ErrorLoggingOTO("OTO-IsHeatThing", userName, ex.Message & ex.Source & " " & dj)
    ''            Return Nothing
    ''        End Try
    ''        Return Nothing
    ''    End Using
    ''End Function


    Public Function UpdateKanbanLabel(ByVal userName As String, ByVal kbCode As String, ByVal subType As String, ByVal qty As Double) As String
        Using da As DataAccess = GetDataAccess()
            Dim msg As String = String.Empty
            Try
                Dim sql As String
                If (Not String.Equals(subType, "PCB", StringComparison.OrdinalIgnoreCase)) Then   '' Not heat thing
                    sql = String.Format("exec SP_UpdateKanbanLabel '{0}', '{1}'", kbCode, qty)
                Else
                    sql = String.Format("exec SP_UpdateKanbanLabel '{0}', '{1}', 'PCB'", kbCode, qty)
                End If
                msg = da.ExecuteScalar(sql, otoConn).ToString()
            Catch ex As Exception
                ErrorLoggingOTO("OTO-UpdateKanbanLabel", userName, ex.Message & ex.Source & kbCode & subType & qty)
                Return "Kanban:" + kbCode + "update failed."
            End Try
            Return msg

        End Using

    End Function

#End Region

    ''' <summary>
    ''' Error log for One to One database.
    ''' </summary>
    ''' <param name="ModuleName"></param>
    ''' <param name="User"></param>
    ''' <param name="ErrMsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ErrorLoggingOTO(ByVal ModuleName As String, ByVal User As String, ByVal ErrMsg As String) As Boolean

        '' Use stored procedure instead of SQL.
        '' Modify by Jackson Huang on 12/02/2013
        Dim ra As Integer
        Dim myConnect As SqlClient.SqlConnection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("eTraceOTOConnString"))
        ''Dim RecordErrorCommand As SqlClient.SqlCommand

        Try
            Using com = New SqlClient.SqlCommand("sp_ErrorLogging", myConnect)
                ' Uses store procedure.
                com.CommandType = CommandType.StoredProcedure

                ' Pass jobId to store procedure.
                com.Parameters.Add("@Module", SqlDbType.Char).Value = ModuleName
                com.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = User
                com.Parameters.Add("@ErrorMsg", SqlDbType.NVarChar).Value = ErrMsg
                com.Connection.Open()
                ra = com.ExecuteNonQuery()
                com.Connection.Close()
            End Using
        Catch ex As Exception
            myConnect.Close()
            Return False
        End Try

        If (ra = -1) Then
            Return True
        End If

        Return False

    End Function


    Public Function EmployeeIDLoginOTO(ByVal EmployeeID As String) As DataSet
        Dim strSQL As String
        Dim dsTmp As New DataSet
        Dim da As DataAccess = GetDataAccess()
        Try
            strSQL = String.Format("select AccessCardID,EmployeeID,Name from T_AccessCard where EmployeeID='" + EmployeeID + "'")
            dsTmp = da.ExecuteDataSet(strSQL)
            Return dsTmp
        Catch ex As Exception
            ErrorLoggingOTO("OTO-EmployeeIDLoginOTO", "", ex.Message)
            Return Nothing
        End Try
    End Function

End Class
